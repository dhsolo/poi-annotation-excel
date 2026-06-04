# Contributing

Thanks for your interest in improving poi-annotation-excel!

## Build & test

Requires JDK 17+ and Maven.

```bash
mvn clean test        # compile + run the test suite (no external services needed)
mvn clean verify      # also builds the sources/javadoc jars
```

The import benchmark is disabled by default; run it explicitly:

```bash
mvn -o test -Dtest=ImportBenchmarkTest -Dexcel.benchmark=true -Dexcel.benchmark.rows=100000
```

## Guidelines

- Keep changes focused; one logical change per pull request.
- Add or update tests for any behavior change — the suite runs fully in-memory.
- Match the surrounding code style (naming, comment density, Javadoc on public API).
- Every source file must carry the Apache-2.0 license header (see existing files).
- Avoid introducing non-ASCII characters in file paths; the Javadoc build is sensitive to them.

## Reporting issues

When filing a bug, include the Excel scenario (export vs import, annotations used),
a minimal reproducing model class, and the POI/JDK versions.

## Security

Image export downloads remote URLs supplied in your data. Treat those URLs as untrusted:
the downloader restricts protocols to http/https and caps response size, but callers are
responsible for validating/whitelisting the hosts they feed in.
