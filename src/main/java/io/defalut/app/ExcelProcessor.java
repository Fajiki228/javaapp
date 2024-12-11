package io.defalut.app;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/**
 * Класс для обработки данных из Excel-файлов и генерации отчетов.
 */
public class ExcelProcessor {
    private static final Logger logger = LogManager.getLogger(ExcelProcessor.class);

    /**
     * Обрабатывает входной Excel-файл, анализирует данные и создает новый файл с отчетом.
     *
     * @param inputFile файл, выбранный пользователем.
     * @throws Exception если данные в файле имеют неверный формат.
     */
    public void processExcelFile(File inputFile) throws Exception {
        logger.info("Processing file: {}", inputFile.getName());

        try (XSSFWorkbook workbook = new XSSFWorkbook(inputFile)) {
            Sheet sheet = workbook.getSheetAt(0);

            if (sheet.getPhysicalNumberOfRows() < 2) {
                throw new Exception("Неверный формат файла: Таблица должна содержать хотя бы одну строку");
            }

            Row headerRow = sheet.getRow(0);
            if (headerRow == null || headerRow.getPhysicalNumberOfCells() < 2) {
                throw new Exception("Неверный формат файла: Таблица должна содержать хотя бы два столбца");
            }

            List<StudentRecord> students = parseStudentData(sheet);

            try (XSSFWorkbook outputWorkbook = new XSSFWorkbook()) {
                XSSFSheet outputSheet = outputWorkbook.createSheet("Сводка");

                CellStyle headerStyle = createHeaderStyle(outputWorkbook);
                CellStyle dataStyle = createDataStyle(outputWorkbook);
                CellStyle averageStyle = createAverageStyle(outputWorkbook);

                String[] headers = {"Имя", "Фамилия", "Отчество", "Оценка"};
                Row headerRow2 = outputSheet.createRow(0);
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow2.createCell(i);
                    cell.setCellValue(headers[i]);
                    cell.setCellStyle(headerStyle);
                    outputSheet.setColumnWidth(i, 15 * 256);
                }

                int rowNum = 1;
                Map<String, Integer> markDistribution = new HashMap<>();
                double totalMarks = 0;
                int validMarksCount = 0;

                for (StudentRecord student : students) {
                    Row row = outputSheet.createRow(rowNum++);

                    Cell firstNameCell = row.createCell(0);
                    firstNameCell.setCellValue(student.firstName);
                    firstNameCell.setCellStyle(dataStyle);

                    Cell middleNameCell = row.createCell(1);
                    middleNameCell.setCellValue(student.middleName);
                    middleNameCell.setCellStyle(dataStyle);

                    Cell lastNameCell = row.createCell(2);
                    lastNameCell.setCellValue(student.lastName);
                    lastNameCell.setCellStyle(dataStyle);

                    String markValue = student.mark >= 3 && student.mark <= 5 ?
                        String.valueOf(student.mark) : "не допущен";

                    Cell markCell = row.createCell(3);
                    markCell.setCellValue(markValue);
                    markCell.setCellStyle(dataStyle);

                    markDistribution.merge(markValue, 1, Integer::sum);
                    if (!markValue.equals("не допущен")) {
                        totalMarks += student.mark;
                        validMarksCount++;
                    }
                }

                Row averageRow = outputSheet.createRow(rowNum + 1);
                Cell averageLabelCell = averageRow.createCell(0);
                averageLabelCell.setCellValue("Средняя оценка");
                averageLabelCell.setCellStyle(averageStyle);

                Cell averageValueCell = averageRow.createCell(1);
                averageValueCell.setCellValue(totalMarks / validMarksCount);
                averageValueCell.setCellStyle(averageStyle);

                createHistogram(outputWorkbook, outputSheet, markDistribution);

                String outputPath = getOutputFilePath(inputFile);
                try (FileOutputStream fileOut = new FileOutputStream(outputPath)) {
                    outputWorkbook.write(fileOut);
                }

                Desktop.getDesktop().open(new File(outputPath));
            }
        }
    }

    /**
     * Создает стиль для заголовков таблицы в Excel-файле.
     *
     * @param workbook текущая рабочая книга Excel.
     * @return стиль, применяемый к ячейкам заголовков.
     */
    private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        return style;
    }

    /**
     * Создает стиль для данных таблицы в Excel-файле.
     *
     * @param workbook текущая рабочая книга Excel.
     * @return стиль, применяемый к ячейкам с данными.
     */
    private CellStyle createDataStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Создает стиль для отображения средней оценки в Excel-файле.
     *
     * @param workbook текущая рабочая книга Excel.
     * @return стиль, применяемый к ячейке со средней оценкой.
     */
    private CellStyle createAverageStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * Извлекает данные о студентах из переданного листа Excel.
     *
     * @param sheet лист Excel, содержащий данные о студентах.
     * @return список записей студентов.
     * @throws Exception если данные в таблице имеют некорректный формат.
     */
    private List<StudentRecord> parseStudentData(Sheet sheet) throws Exception {
        List<StudentRecord> students = new ArrayList<>();

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;

            Cell nameCell = row.getCell(0);
            Cell markCell = row.getCell(1);

            if (nameCell == null || markCell == null) {
                throw new Exception("Неверный формат файла: Отсутствует информация");
            }

            if (markCell.getCellType() != CellType.NUMERIC) {
                throw new Exception("Неверный формат файла: Оценка должна быть числом");
            }

            String fullName = nameCell.getStringCellValue().trim();
            if (fullName.isEmpty()) {
                throw new Exception("Неверный формат файла: Пустое имя в ряду");
            }

            double mark = markCell.getNumericCellValue();
            if (mark < 0) {
                throw new Exception("Неверный формат файла: Оценка не может быть негативной");
            }

            String[] nameParts = fullName.split("\\s+");
            if (nameParts.length < 2) {
                throw new Exception("Неверный формат файла: В имени должно содержаться хотя бы имя и фамилия");
            }

            String firstName = nameParts[0];
            String middleName = nameParts.length > 2 ? nameParts[1] : "";
            String lastName = nameParts[nameParts.length - 1];

            students.add(new StudentRecord(firstName, middleName, lastName, mark));
        }

        return students;
    }

    /**
     * Создает диаграмму распределения оценок в Excel-файле.
     *
     * @param workbook рабочая книга, в которой будет создан график.
     * @param sheet лист, на котором создается гистограмма.
     * @param distribution распределение оценок (оценка -> количество студентов).
     */
    private void createHistogram(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, Integer> distribution) {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 6, 1, 15, 15);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("График оценок");
        chart.setTitleOverlay(false);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Оценки");
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("Студенты");

        String[] marks = distribution.keySet().toArray(new String[0]);
        Integer[] counts = distribution.values().toArray(new Integer[0]);

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromArray(marks);
        XDDFNumericalDataSource<Integer> values = XDDFDataSourcesFactory.fromArray(counts);

        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        XDDFChartData.Series series = data.addSeries(categories, values);
        series.setTitle("Студенты", null);

        chart.plot(data);
    }

    /**
     * Формирует путь для сохранения выходного Excel-файла на основе имени входного файла.
     *
     * @param inputFile входной файл Excel.
     * @return путь для сохранения выходного файла.
     */
    private String getOutputFilePath(File inputFile) {
        String path = inputFile.getAbsolutePath();
        return path.substring(0, path.lastIndexOf(".")) + "_info.xlsx";
    }
}

/**
 * Класс для хранения данных о студенте.
 */
class StudentRecord {
    String firstName;
    String middleName;
    String lastName;
    double mark;

    public StudentRecord(String firstName, String middleName, String lastName, double mark) {
        this.firstName = firstName;
        this.middleName = middleName;
        this.lastName = lastName;
        this.mark = mark;
    }
}
