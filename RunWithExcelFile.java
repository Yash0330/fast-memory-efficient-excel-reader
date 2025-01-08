import java.util.concurrent.atomic.AtomicInteger;

public class RunWithExcelFile {
    public static void main(String[] args) throws InterruptedException {
        FastExcelReader reader = new FastExcelReader();
        String filePath = "C:\\Users\\yashw\\Downloads\\fast-memory-efficient-excel-reader\\my_data.xlsx";
        Thread.sleep(10 * 1000);

        long startTime = System.nanoTime();
        try {
            // Read the Excel file and get a stream of rows
            AtomicInteger i = new AtomicInteger();
            reader.readExcel(filePath)
                    .forEach(row -> {
                        // Process each row
                        System.out.println("Row Data:");
                        row.forEach((columnName, value) -> {
                            System.out.println("Column: " + columnName + ", Value: " + value);
                        });
                        System.out.println("---- End of Row ----");
                        i.getAndIncrement();
                    });
            long endTime = System.nanoTime();
            System.out.println("Successfully processed " + i.get() + " rows in " + (endTime - startTime) / 1e9 + " seconds.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}