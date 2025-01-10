import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamReader;
import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FilterInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * FastExcelReader is a lightweight and efficient utility for reading Excel files (XLSX format).
 * It uses streaming to process large files without consuming excessive memory.
 *
 * <p>Key features: <br>
 * - Handles shared strings for Excel. <br>
 * - Supports configurable buffered input stream size. <br>
 * - Reads sheet data row by row. <br>
 * - Allows specifying a subset of column names to be read. <br>
 *
 * <p>Note: For parallel execution, create separate instances of FastExcelReader for each thread
 * to avoid shared state issues. This ensures thread safety while processing multiple files concurrently.
 *
 * <p>Usage Example:
 * <pre>
 * {@code
 * FastExcelReader reader = new FastExcelReader(4096, List.of("Name", "Age", "Country"));
 * try (Stream<Map<String, String>> rows = reader.readExcel("path/to/excel.xlsx")) {
 *     rows.forEach(row -> {
 *         System.out.println("Row: " + row);
 *     });
 * }
 * }
 * </pre>
 * <p>
 * Author: Yashwanth M
 */
public class FastExcelReader {

    // XML element constants
    private static final String ELEMENT_V = "v";
    private static final String ELEMENT_SI = "si";
    private static final String ELEMENT_T = "t";
    private static final String ELEMENT_C = "c";
    private static final String ELEMENT_ROW = "row";

    private final Map<Integer, String> columnMap = new HashMap<>();
    private final List<String> sharedStrings = new ArrayList<>();
    private List<String> columnNames = null;
    private int bufferedInputStreamSize = 2048;

    /**
     * Default constructor with default buffer size.
     */
    public FastExcelReader() {
    }

    /**
     * Constructor with configurable buffer size.
     *
     * @param bufferedInputStreamSize size of the buffer for reading the input stream.
     */
    public FastExcelReader(int bufferedInputStreamSize) {
        this.bufferedInputStreamSize = bufferedInputStreamSize;
    }

    /**
     * Constructor with specific column names to filter.
     *
     * @param columnNames list of column names to be read.
     */
    public FastExcelReader(List<String> columnNames) {
        this.columnNames = columnNames;
    }

    /**
     * Constructor with both buffer size and column names.
     *
     * @param bufferedInputStreamSize size of the buffer for reading the input stream.
     * @param columnNames             list of column names to be read.
     */
    public FastExcelReader(int bufferedInputStreamSize, List<String> columnNames) {
        this.bufferedInputStreamSize = bufferedInputStreamSize;
        this.columnNames = columnNames;
    }

    /**
     * Reads an Excel file and returns its data as a stream of maps, where each map represents a row.
     *
     * @param filePath path to the Excel file.
     * @return a stream of rows, each represented as a map of column names to values.
     * @throws Exception if an error occurs during file reading or parsing.
     */
    public Stream<Map<String, String>> readExcel(String filePath) throws Exception {
        try (FileInputStream fis = new FileInputStream(filePath);
             BufferedInputStream bis = new BufferedInputStream(fis, bufferedInputStreamSize);
             ZipInputStream zis = new ZipInputStream(bis)) {

            ZipEntry entry;

            // First pass: Process shared strings
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().equals("xl/sharedStrings.xml")) {
                    parseSharedStrings(zis);
                    break;
                }
            }
        }

        // Re-open the stream for the second pass
        try (FileInputStream fis = new FileInputStream(filePath);
             BufferedInputStream bis = new BufferedInputStream(fis, bufferedInputStreamSize);
             ZipInputStream zis = new ZipInputStream(bis)) {

            ZipEntry entry;

            // Second pass: Process sheet data
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().startsWith("xl/worksheets/sheet")) {
                    return parseSheetData(zis);
                }
            }
        }

        return Stream.empty();
    }

    /**
     * Parses shared strings XML and stores them in a list.
     *
     * @param stream input stream of the shared strings XML.
     * @throws Exception if an error occurs during parsing.
     */
    private void parseSharedStrings(InputStream stream) throws Exception {
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLStreamReader reader = factory.createXMLStreamReader(stream);

        while (reader.hasNext()) {
            int event = reader.next();

            if (event == XMLStreamConstants.START_ELEMENT && ELEMENT_SI.equals(reader.getLocalName())) {
                sharedStrings.add(parseStringItem(reader));
            }
        }
    }

    /**
     * Parses a string item from the shared strings XML.
     *
     * @param reader XML stream reader positioned at a shared string item.
     * @return the parsed string value.
     * @throws Exception if an error occurs during parsing.
     */
    private String parseStringItem(XMLStreamReader reader) throws Exception {
        StringBuilder sb = new StringBuilder();

        while (reader.hasNext()) {
            int event = reader.next();

            if (event == XMLStreamConstants.START_ELEMENT && ELEMENT_T.equals(reader.getLocalName())) {
                sb.append(reader.getElementText());
            } else if (event == XMLStreamConstants.END_ELEMENT && ELEMENT_SI.equals(reader.getLocalName())) {
                break;
            }
        }

        return sb.toString();
    }

    /**
     * Parses sheet data into a stream of row maps.
     *
     * @param stream input stream of the sheet XML.
     * @return a stream of rows, each represented as a map of column names to values.
     * @throws Exception if an error occurs during parsing.
     */
    private Stream<Map<String, String>> parseSheetData(InputStream stream) throws Exception {
        InputStream nonClosingStream = new NonClosingInputStream(stream);
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLStreamReader reader = factory.createXMLStreamReader(nonClosingStream);
        boolean isHeaderRow = true;

        List<Map<String, String>> rows = new ArrayList<>();

        try {
            while (reader.hasNext()) {
                int event = reader.next();

                if (event == XMLStreamConstants.START_ELEMENT && ELEMENT_ROW.equals(reader.getLocalName())) {
                    if (isHeaderRow) {
                        parseHeaderRow(reader);
                        isHeaderRow = false;
                    } else {
                        Map<String, String> rowValues = processRow(reader);
                        rows.add(rowValues);
                    }
                }
            }
        } finally {
            reader.close();
        }

        return rows.stream();
    }

    /**
     * Parses the header row to map column indices to their respective column names.
     * The column map is used to match cells with their corresponding column names.
     *
     * <p>Note:
     * - Columns are added only if no filter is provided, or if they match the specified filter.
     *
     * @param reader XMLStreamReader instance to parse the header row.
     * @throws Exception if an error occurs while reading the XML data.
     */
    private void parseHeaderRow(XMLStreamReader reader) throws Exception {
        int colIndex = 0;

        while (reader.hasNext()) {
            int event = reader.next();

            if (event == XMLStreamConstants.START_ELEMENT && ELEMENT_C.equals(reader.getLocalName())) {
                String headerValue = processCell(reader);

                if (headerValue != null && !headerValue.trim().isEmpty()) {
                    // Add the column if no filter is provided, or if it matches the filter
                    if (columnNames == null || columnNames.contains(headerValue)) {
                        columnMap.put(colIndex, headerValue);
                    }
                }

                colIndex++;
            } else if (event == XMLStreamConstants.END_ELEMENT && ELEMENT_ROW.equals(reader.getLocalName())) {
                break;
            }
        }
    }

    /**
     * Processes a single data row and maps cell values to their respective column names.
     * Returns a map where the keys are column names and values are the corresponding cell values.
     *
     * <p>Note:
     * - Adds null values for missing columns to ensure the map is consistent with the header.
     *
     * @param reader XMLStreamReader instance to parse the row.
     * @return a map representing the row with column names as keys and cell values as values.
     * @throws Exception if an error occurs while reading the XML data.
     */
    private Map<String, String> processRow(XMLStreamReader reader) throws Exception {
        Map<String, String> rowValues = new HashMap<>();

        while (reader.hasNext()) {
            int event = reader.next();

            if (event == XMLStreamConstants.START_ELEMENT && ELEMENT_C.equals(reader.getLocalName())) {
                String cellRef = reader.getAttributeValue(null, "r");
                int colIndex = getColumnIndex(cellRef);

                // Process only valid columns in the column map
                if (columnMap.containsKey(colIndex)) {
                    String columnName = columnMap.get(colIndex);
                    String cellValue = processCell(reader);
                    rowValues.put(columnName, cellValue);
                }
            } else if (event == XMLStreamConstants.END_ELEMENT && ELEMENT_ROW.equals(reader.getLocalName())) {
                break;
            }
        }

        // Ensure all columns have a value, even if null
        for (String columnName : columnMap.values()) {
            rowValues.putIfAbsent(columnName, null);
        }

        return rowValues;
    }

    /**
     * Processes an individual cell to extract its value.
     * Supports both shared strings and direct values.
     *
     * @param reader XMLStreamReader instance to parse the cell.
     * @return the cell value as a string, or null if the cell is empty.
     * @throws Exception if an error occurs while reading the XML data.
     */
    private String processCell(XMLStreamReader reader) throws Exception {
        String cellValue = null;
        String type = reader.getAttributeValue(null, "t");

        while (reader.hasNext()) {
            int event = reader.next();

            if (event == XMLStreamConstants.START_ELEMENT && ELEMENT_V.equals(reader.getLocalName())) {
                String rawValue = reader.getElementText();
                if ("s".equals(type)) {
                    cellValue = sharedStrings.get(Integer.parseInt(rawValue));
                } else {
                    cellValue = rawValue;
                }
            } else if (event == XMLStreamConstants.END_ELEMENT && ELEMENT_C.equals(reader.getLocalName())) {
                break;
            }
        }

        return cellValue;
    }

    /**
     * Converts an Excel cell reference (e.g., "A1") to a zero-based column index.
     *
     * @param cellRef the cell reference (e.g., "A1").
     * @return the zero-based column index.
     */
    private int getColumnIndex(String cellRef) {
        int colIndex = 0;
        for (int i = 0; i < cellRef.length(); i++) {
            char ch = cellRef.charAt(i);
            if (Character.isDigit(ch)) {
                break; // Stop when digits are encountered
            }
            colIndex = colIndex * 26 + (ch - 'A' + 1);
        }
        return colIndex - 1; // Convert to zero-based index
    }
}

/**
 * A custom InputStream wrapper that prevents the underlying stream from being closed.
 * This is useful for wrapping streams that need to remain open for further processing.
 */
class NonClosingInputStream extends FilterInputStream {
    /**
     * Constructs a NonClosingInputStream.
     *
     * @param in the input stream to wrap.
     */
    protected NonClosingInputStream(InputStream in) {
        super(in);
    }

    /**
     * Overrides the close method to do nothing, ensuring the wrapped stream remains open.
     */
    @Override
    public void close() {
        // Prevents the underlying InputStream from being closed
    }
}
