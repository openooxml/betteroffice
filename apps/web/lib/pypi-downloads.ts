import { PYPI_DISTRIBUTIONS } from "../../../scripts/python-binding-names.mjs";

/** PyPI has no org listing endpoint, so the binding registry is the source. */
export const PYPI_PACKAGES: string[] = PYPI_DISTRIBUTIONS;
