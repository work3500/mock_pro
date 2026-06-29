# AGENTS.md

## Cursor Cloud specific instructions

This is a minimal Maven Java CLI application (`org.example:mock_pro`). The only runtime
artifact is `org.example.App`, which prints `Hello World!` to stdout. There are no
servers, databases, or external services to run.

- Toolchain: JDK 21 (preinstalled) + Maven (installed by the update script). There is no
  `mvnw` wrapper, so the system `mvn` is used.
- Build/test/run caveat: the `pom.xml` does not set a Java compiler release, so Maven's
  default `maven-compiler-plugin` falls back to source/target `5`, which JDK 21 rejects
  (`Source option 5 is no longer supported`). Always pass an explicit release when
  building/testing:
  - Build + test: `mvn clean package -Dmaven.compiler.source=8 -Dmaven.compiler.target=8`
  - Test only: `mvn test -Dmaven.compiler.source=8 -Dmaven.compiler.target=8`
  - Run: `java -cp target/classes org.example.App` (after a successful build)
- Lint: no linter is configured in `pom.xml` (no Checkstyle/SpotBugs/PMD/Spotless).
