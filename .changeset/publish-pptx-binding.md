---
"@betteroffice/python-pptx": patch
---

`betteroffice-pptx` publishes to PyPI. The distribution joins the release train's publish matrix, and a `workflow_dispatch` workflow can publish one binding on its own — to create a PyPI project from a pending Trusted Publisher, or to fill in a wheel a release run dropped. Publishing is Trusted Publishing only; the workflow refuses to run while a repository-scoped `PYPI_API_TOKEN` exists.
