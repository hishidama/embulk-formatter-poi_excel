package org.embulk.formatter.poi_excel;

import com.google.common.base.Optional;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.embulk.config.Config;
import org.embulk.config.ConfigDefault;
import org.embulk.config.ConfigInject;
import org.embulk.config.ConfigSource;
import org.embulk.config.Task;
import org.embulk.config.TaskSource;
import org.embulk.formatter.poi_excel.visitor.PoiExcelColumnVisitor;
import org.embulk.spi.BufferAllocator;
import org.embulk.spi.FileOutput;
import org.embulk.spi.FormatterPlugin;
import org.embulk.spi.Page;
import org.embulk.spi.PageOutput;
import org.embulk.spi.PageReader;
import org.embulk.spi.Schema;
import org.embulk.spi.time.TimestampFormatter;
import org.embulk.spi.time.TimestampFormatter.TimestampColumnOption;
import org.embulk.spi.util.FileOutputOutputStream;
import org.embulk.spi.util.FileOutputOutputStream.CloseMode;

import java.io.IOException;
import java.text.MessageFormat;
import java.util.Map;

public class PoiExcelFormatterPlugin
        implements FormatterPlugin
{
    public interface PluginTask
            extends Task, TimestampFormatter.Task
    {
        @Config("spread_sheet_version")
        @ConfigDefault("\"EXCEL2007\"")
        public SpreadsheetVersion getSpreadsheetVersion();

        @Config("sheet_name")
        @ConfigDefault("\"Sheet1\"")
        public String getSheetName();

        @Config("column_options")
        @ConfigDefault("{}")
        public Map<String, ColumnOption> getColumnOptions();

        @ConfigInject
        public BufferAllocator getBufferAllocator();
    }

    public interface ColumnOption
            extends Task, TimestampColumnOption
    {

        @Config("data_format")
        @ConfigDefault("null")
        public Optional<String> getDataFormat();
    }

    @Override
    public void transaction(ConfigSource config, Schema schema, FormatterPlugin.Control control)
    {
        PluginTask task = config.loadConfig(PluginTask.class);

        control.run(task.dump());
    }

    @Override
    public PageOutput open(TaskSource taskSource, final Schema schema, FileOutput output)
    {
        final PluginTask task = taskSource.loadTask(PluginTask.class);

        final Sheet sheet = newWorkbook(task);

        final FileOutputOutputStream stream = new FileOutputOutputStream(output, task.getBufferAllocator(),
                CloseMode.CLOSE);
        stream.nextFile();

        return new PageOutput()
        {
            private final PageReader pageReader = new PageReader(schema);

            @Override
            public void add(Page page)
            {
                pageReader.setPage(page);
                PoiExcelColumnVisitor visitor = new PoiExcelColumnVisitor(task, schema, sheet, pageReader);
                while (pageReader.nextRecord()) {
                    schema.visitColumns(visitor);
                    visitor.endRecord();
                }
            }

            @Override
            public void finish()
            {
                Workbook book = sheet.getWorkbook();
                try (FileOutputOutputStream os = stream) {
                    book.write(os);
                    os.finish();
                }
                catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }

            @Override
            public void close()
            {
                stream.close();
            }
        };
    }

    @SuppressWarnings("resource")
    protected Sheet newWorkbook(PluginTask task)
    {
        Workbook book;
        {
            SpreadsheetVersion version = task.getSpreadsheetVersion();
            switch (version) {
                case EXCEL97:
                    book = new HSSFWorkbook();
                    break;
                case EXCEL2007:
                    book = new XSSFWorkbook();
                    break;
                default:
                    throw new UnsupportedOperationException(MessageFormat.format("unsupported spread_sheet_version={0}",
                            version));
            }
        }

        String sheetName = task.getSheetName();
        return book.createSheet(sheetName);
    }
}
