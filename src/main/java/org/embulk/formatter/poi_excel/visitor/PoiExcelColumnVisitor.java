package org.embulk.formatter.poi_excel.visitor;

import com.google.common.base.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.embulk.formatter.poi_excel.PoiExcelFormatterPlugin.ColumnOption;
import org.embulk.formatter.poi_excel.PoiExcelFormatterPlugin.PluginTask;
import org.embulk.spi.Column;
import org.embulk.spi.ColumnVisitor;
import org.embulk.spi.PageReader;
import org.embulk.spi.Schema;
import org.embulk.spi.time.Timestamp;
import org.embulk.spi.time.TimestampFormatter;
import org.embulk.spi.util.Timestamps;
import org.joda.time.DateTime;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class PoiExcelColumnVisitor
        implements ColumnVisitor
{

    private final PluginTask task;
    private final Schema schema;
    private final Sheet sheet;
    private final PageReader pageReader;

    private int rowIndex = 0;

    private Row currentRow = null;

    public PoiExcelColumnVisitor(PluginTask task, Schema schema, Sheet sheet, PageReader pageReader)
    {
        this.task = task;
        this.schema = schema;
        this.sheet = sheet;
        this.pageReader = pageReader;
    }

    @Override
    public void booleanColumn(Column column)
    {
        if (pageReader.isNull(column)) {
            return;
        }
        boolean value = pageReader.getBoolean(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void longColumn(Column column)
    {
        if (pageReader.isNull(column)) {
            return;
        }
        long value = pageReader.getLong(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void doubleColumn(Column column)
    {
        if (pageReader.isNull(column)) {
            return;
        }
        double value = pageReader.getDouble(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void stringColumn(Column column)
    {
        if (pageReader.isNull(column)) {
            return;
        }
        String value = pageReader.getString(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void jsonColumn(Column column)
    {
        throw new UnsupportedOperationException("This plugin doesn't support json type. Please try to upgrade version of the plugin using 'embulk gem update' command. If the latest version still doesn't support json type, please contact plugin developers, or change configuration of input plugin not to use json type.");
    }

    @Override
    public void timestampColumn(Column column)
    {
        if (pageReader.isNull(column)) {
            return;
        }
        TimestampFormatter formatter = getTimestampFormatter(column);
        Timestamp timestamp = pageReader.getTimestamp(column);
        DateTime dateTime = new DateTime(timestamp.toEpochMilli(), formatter.getTimeZone());
        Date value = dateTime.toDate();
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    private TimestampFormatter[] timestampFormatters;

    protected final TimestampFormatter getTimestampFormatter(Column column)
    {
        if (timestampFormatters == null) {
            timestampFormatters = Timestamps.newTimestampColumnFormatters(task, schema, task.getColumnOptions());
        }
        return timestampFormatters[column.getIndex()];
    }

    protected Cell getCell(Column column)
    {
        Cell cell = CellUtil.getCell(getRow(), column.getIndex());

        ColumnOption option = getColumnOption(column);
        if (option != null) {
            Optional<String> formatOption = option.getDataFormat();
            if (formatOption.isPresent()) {
                String formatString = formatOption.get();
                CellStyle style = styleMap.get(formatString);
                if (style == null) {
                    Workbook book = sheet.getWorkbook();
                    style = book.createCellStyle();
                    CreationHelper helper = book.getCreationHelper();
                    short fmt = helper.createDataFormat().getFormat(formatString);
                    style.setDataFormat(fmt);
                    styleMap.put(formatString, style);
                }
                cell.setCellStyle(style);
            }
        }

        return cell;
    }

    protected final ColumnOption getColumnOption(Column column)
    {
        Map<String, ColumnOption> map = task.getColumnOptions();
        return map.get(column.getName());
    }

    private Map<String, CellStyle> styleMap = new HashMap<>();

    private Row getRow()
    {
        if (currentRow == null) {
            currentRow = sheet.createRow(rowIndex);
        }
        return currentRow;
    }

    public void endRecord()
    {
        rowIndex++;
        currentRow = null;
    }
}
