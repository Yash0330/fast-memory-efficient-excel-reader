# fast-memory-efficient-excel-reader

`fast-memory-efficient-excel-reader` is a lightweight and efficient Java library for reading large Excel files (`.xlsx`) with minimal memory usage. It leverages streaming to process data row by row, making it ideal for scenarios where memory consumption is a concern.

---

## Features

- **Memory-Efficient**: Processes data row-by-row using streaming, avoiding loading entire files into memory.
- **Fast Processing**: Handles large Excel files with high performance.
- **Custom Column Selection**: Supports selecting specific columns to process.
- **Thread-Safe**: Parallel execution is possible by creating separate instances of the `FastExcelReader` class.
- **Simple API**: Easy-to-use methods for reading Excel data as a stream of maps.

---

## Usage

### Basic Example

The library reads `.xlsx` files and returns data as a stream of `Map<String, String>` where keys are column names, and values are cell values.

```java
import java.util.Map;
import java.util.stream.Stream;

public class Example {
    public static void main(String[] args) {
        try {
            FastExcelReader reader = new FastExcelReader();
            Stream<Map<String, String>> rows = reader.readExcel("example.xlsx");

            rows.forEach(row -> {
                System.out.println(row);
            });
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
```
### Using Column Filters
To process only specific columns, provide a list of column names:

```java
import java.util.Arrays;
import java.util.Map;
import java.util.stream.Stream;

public class ExampleWithFilters {
    public static void main(String[] args) {
        try {
            FastExcelReader reader = new FastExcelReader(Arrays.asList("Name", "Age"));
            Stream<Map<String, String>> rows = reader.readExcel("example.xlsx");

            rows.forEach(row -> {
                System.out.println(row);
            });
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
```
### Parallel Execution
To process multiple files in parallel, create separate instances of FastExcelReader:

``` java
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ParallelExample {
    public static void main(String[] args) {
        ExecutorService executor = Executors.newFixedThreadPool(2);

        executor.execute(() -> {
            try {
                FastExcelReader reader1 = new FastExcelReader();
                reader1.readExcel("file1.xlsx").forEach(System.out::println);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });

        executor.execute(() -> {
            try {
                FastExcelReader reader2 = new FastExcelReader();
                reader2.readExcel("file2.xlsx").forEach(System.out::println);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });

        executor.shutdown();
    }
}
```
---

## Key Methods

### `readExcel(String filePath)`
Reads an Excel file and returns a `Stream<Map<String, String>>`. Each map represents a row, where:
- The **key** is the column name.
- The **value** is the corresponding cell value for that row.

### Constructors
- `FastExcelReader()`: Default constructor.
- `FastExcelReader(List<String> columnNames)`: Reads only specified columns.
- `FastExcelReader(int bufferedInputStreamSize)`: Custom buffer size.
- `FastExcelReader(int bufferedInputStreamSize, List<String> columnNames)`: Combines buffer size and column selection.

### Reference Examples

For reference examples on how to use this library with an Excel file, you can refer to the `RunWithExcelFile` class available in the project.

## Best Practices

- **Thread Safety**: Ensure separate instances of `FastExcelReader` for each thread or task.
- **Resource Management**: Use `try-with-resources` or close streams properly to avoid resource leaks.
- **Large Files**: Test the buffer size (`bufferedInputStreamSize`) for optimal performance when handling very large files.

## Limitations

- Supports only `.xlsx` files.
- Does not support advanced Excel features like formulas, macros, or charts.

## License

This project is licensed under the Apache-2.0 License. See the [LICENSE](https://github.com/Yash0330/fast-memory-efficient-excel-reader/blob/main/LICENSE) file for details.

## Author

**Yashwanth M**  
Connect with me on [LinkedIn](https://www.linkedin.com/in/yash0330) for queries or suggestions.
