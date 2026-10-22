import argparse
import itertools
import json
import math
import random
import statistics
from pathlib import Path

RESAMPLES = 10000
SEED = 0
Z = 1.959963984540054


def percentile(values, fraction):
    index = (len(values) - 1) * fraction
    lower = math.floor(index)
    return values[lower] + (values[math.ceil(index)] - values[lower]) * (index - lower)


def bootstrap(clusters):
    generator = random.Random(SEED)
    estimates = []
    for _ in range(RESAMPLES):
        sample = [clusters[generator.randrange(len(clusters))] for _ in clusters]
        estimates.append(sum(value for value, _ in sample) / sum(weight for _, weight in sample))
    estimates.sort()
    return [percentile(estimates, 0.025), percentile(estimates, 0.975)]


def wilson(passed, total):
    proportion = passed / total
    denominator = 1 + Z * Z / total
    center = (proportion + Z * Z / (2 * total)) / denominator
    radius = Z * math.sqrt(proportion * (1 - proportion) / total + Z * Z / (4 * total * total)) / denominator
    return [0.0 if passed == 0 else max(0.0, center - radius),
            1.0 if passed == total else min(1.0, center + radius)]


def numeric(value, name):
    if value is not None and (isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0):
        raise ValueError(f'invalid {name}: {value!r}')
    return value


def aggregate(paths):
    grouped = {}
    for path in paths:
        report = json.loads(Path(path).read_text())
        arm = report['arm']
        key = (report.get('suite', 'v2'), arm, report['model'], report.get('agent', 'none' if arm.startswith('scripted-') else 'codex'))
        records = grouped.setdefault(key, {})
        for record in report['results']:
            index = record.get('run', 1)
            if type(index) is not int or index < 1:
                raise ValueError('run must be a positive integer')
            pair = (record['task'], index)
            if pair in records:
                raise ValueError(f'duplicate task/run: {key} {pair}')
            passed = record.get('passed', record.get('pass'))
            if type(passed) is not bool or ('pass' in record and record['pass'] != passed):
                raise ValueError(f'invalid pass result: {pair}')
            records[pair] = {'passed': passed, 'tool_calls': numeric(record.get('tool_calls'), 'tool_calls'),
                             'seconds': numeric(record.get('seconds', record.get('duration_seconds')), 'seconds')}
    groups = []
    for (suite, arm, model, agent), records in sorted(grouped.items()):
        if not records:
            raise ValueError(f'no finished runs for {suite}/{arm}/{model}/{agent}')
        task_ids = sorted({task for task, _ in records})
        per_task = []
        for task in task_ids:
            trials = [record for (identity, _), record in sorted(records.items()) if identity == task]
            passed = sum(record['passed'] for record in trials)
            rate = passed / len(trials)
            per_task.append({'task': task, 'runs': len(trials), 'pass_count': passed,
                             'pass_rate': rate, 'variance': rate * (1 - rate)})
        counts = {task['runs'] for task in per_task}
        passed = sum(record['passed'] for record in records.values())
        calls = [record['tool_calls'] for record in records.values() if record['tool_calls'] is not None]
        seconds = [record['seconds'] for record in records.values() if record['seconds'] is not None]
        groups.append({'suite': suite, 'arm': arm, 'model': model, 'agent': agent,
                       'tasks': len(task_ids), 'runs': next(iter(counts)) if len(counts) == 1 else None,
                       'task_runs': len(records), 'pass_count': passed, 'pass_rate': passed / len(records),
                       'bootstrap_95': bootstrap([(task['pass_count'], task['runs']) for task in per_task]),
                       'wilson_95': wilson(passed, len(records)),
                       'all_runs_pass_rate': sum(task['pass_count'] == task['runs'] for task in per_task) / len(task_ids),
                       'any_pass_rate': sum(task['pass_count'] > 0 for task in per_task) / len(task_ids),
                       'per_task': per_task, 'median_tool_calls': statistics.median(calls) if calls else None,
                       'median_seconds': statistics.median(seconds) if seconds else None})
    comparisons = []
    for left, right in itertools.combinations(groups, 2):
        if left['suite'] != right['suite'] or left['model'] != right['model'] or left['arm'] == right['arm']:
            continue
        if [task['task'] for task in left['per_task']] != [task['task'] for task in right['per_task']]:
            continue
        differences = [{'task': a['task'], 'difference': a['pass_rate'] - b['pass_rate']}
                       for a, b in zip(left['per_task'], right['per_task'])]
        comparisons.append({'suite': left['suite'], 'model': left['model'],
                            'arm_a': left['arm'], 'agent_a': left['agent'], 'arm_b': right['arm'], 'agent_b': right['agent'],
                            'tasks': len(differences), 'difference': statistics.mean(item['difference'] for item in differences),
                            'bootstrap_95': bootstrap([(item['difference'], 1) for item in differences]),
                            'per_task': differences})
    return {'confidence_level': 0.95, 'bootstrap_resamples': RESAMPLES, 'bootstrap_seed': SEED,
            'groups': groups, 'comparisons': comparisons}


def markdown(summary):
    def interval(bounds):
        return f'[{bounds[0]:.3f}, {bounds[1]:.3f}]'

    def cell(value):
        return str(value).replace('|', '\\|').replace('\n', ' ')

    rows = ['| Suite | Arm | Model | Agent | Tasks | Runs/task | Pass rate | Cluster 95% | Wilson 95% | All pass | Any pass | Median calls | Median seconds |',
            '| --- | --- | --- | --- | ---: | ---: | ---: | --- | --- | ---: | ---: | ---: | ---: |']
    for group in summary['groups']:
        values = [group[name] for name in ('suite', 'arm', 'model', 'agent', 'tasks')]
        values += [group['runs'] if group['runs'] is not None else 'mixed', f'{group["pass_rate"]:.3f}',
                   interval(group['bootstrap_95']), interval(group['wilson_95']),
                   f'{group["all_runs_pass_rate"]:.3f}', f'{group["any_pass_rate"]:.3f}',
                   group['median_tool_calls'], group['median_seconds']]
        rows.append('| ' + ' | '.join(cell(value) if value is not None else 'n/a' for value in values) + ' |')
    if summary['comparisons']:
        rows += ['', '| Suite | Model | Arm A (agent) | Arm B (agent) | Tasks | A minus B | Paired cluster 95% |',
                 '| --- | --- | --- | --- | ---: | ---: | --- |']
        for pair in summary['comparisons']:
            values = [pair['suite'], pair['model'], f'{pair["arm_a"]} ({pair["agent_a"]})',
                      f'{pair["arm_b"]} ({pair["agent_b"]})', pair['tasks'], f'{pair["difference"]:.3f}', interval(pair['bootstrap_95'])]
            rows.append('| ' + ' | '.join(cell(value) for value in values) + ' |')
    return '\n'.join(rows) + '\n'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('results', nargs='+', type=Path)
    parser.add_argument('--out', required=True, type=Path)
    parser.add_argument('--md', type=Path)
    args = parser.parse_args()
    try:
        summary = aggregate(args.results)
        args.out.write_text(json.dumps(summary, indent=2, allow_nan=False) + '\n')
        if args.md:
            args.md.write_text(markdown(summary))
    except (ValueError, OSError, KeyError) as error:
        parser.error(str(error))


if __name__ == '__main__':
    main()
