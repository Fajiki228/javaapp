package io.defalut.app;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import java.io.File;


/**
 * Основной класс приложения.
 */
public class MyApp extends Application {
    private static final Logger logger = LogManager.getLogger(MyApp.class);
    private Label errorLabel;

    /**
     * Инициализирует главное окно приложения.
     *
     * @param stage главная сцена приложения.
     */
    @Override
    public void start(Stage stage) {
        try {
            Button selectFileButton = new Button("Выбрать файл");
            errorLabel = new Label();
            errorLabel.setTextFill(Color.RED);
            errorLabel.setWrapText(true);

            VBox root = new VBox(10);
            root.setPadding(new Insets(10));
            root.getChildren().addAll(selectFileButton, errorLabel);

            selectFileButton.setOnAction(e -> handleFileSelection(stage));

            Scene scene = new Scene(root, 300, 100);
            stage.setTitle("Анализ успеваемости");
            stage.setScene(scene);
            stage.show();

        } catch (Exception e) {
            logger.error("Произошла критическая ошибка", e);
        }
    }

    /**
     * Обрабатывает выбор файла через диалоговое окно.
     *
     * @param stage текущая сцена для привязки диалога выбора файла.
     */
    private void handleFileSelection(Stage stage) {
        errorLabel.setText("");
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("Excel Files", "*.xlsx")
        );
        File selectedFile = fileChooser.showOpenDialog(stage);
        if (selectedFile != null) {
            try {
                ExcelProcessor processor = new ExcelProcessor();
                processor.processExcelFile(selectedFile);
            } catch (Exception e) {
                logger.error("Ошибка обработки файла Excel", e);
                String errorMessage = getErrorMessage(e);
                errorLabel.setText(errorMessage);
            }
        }
    }

    /**
     * Возвращает сообщение об ошибке в зависимости от типа исключения.
     *
     * @param e исключение, возникшее при обработке файла.
     * @return строка с описанием ошибки.
     */
    private String getErrorMessage(Exception e) {
        if (e.getCause() instanceof IllegalStateException &&
            e.getCause().getMessage().contains("Cannot get a NUMERIC value from a STRING cell")) {
            return "Неверный формат файла: Столбец оценок должен содержать только числа";
        } else if (e.getMessage().contains("Invalid file format")) {
            return e.getMessage();
        }
        return "Неверный формат файла, необходим действительный xlsx";
    }

    /**
     * Точка входа в приложение.
     *
     * @param args аргументы командной строки.
     */
    public static void main(String[] args) {
        launch();
    }
}
