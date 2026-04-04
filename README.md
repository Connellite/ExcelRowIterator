[![Build](https://github.com/connellite/ExcelRowIterator/actions/workflows/ci.yml/badge.svg)](https://github.com/connellite/ExcelRowIterator/actions/workflows/ci.yml)
[![Maven Central Version](https://img.shields.io/maven-central/v/io.github.connellite/ExcelRowIterator)](https://mvnrepository.com/artifact/io.github.connellite/ExcelRowIterator)

# ExcelRowIterator

Small Java 17 library: forward-only iterators and streams over Apache POI sheets as maps (typed `Object` values or plain strings). Implements `Iterable` and `AutoCloseable` in the same spirit as [ExtraLib](https://github.com/connellite/ExtraLib) JDBC row iterators.

## Requirements

- JDK 17+
- Maven 3.9+

## Dependency

```xml
<dependency>
    <groupId>io.github.connellite</groupId>
    <artifactId>ExcelRowIterator</artifactId>
    <version>0.1.0</version>
</dependency>
```


This artifact depends only on `poi` (usermodel). To open `.xlsx` workbooks, add `poi-ooxml` in your project.

## License

MIT — see [LICENSE](LICENSE).
