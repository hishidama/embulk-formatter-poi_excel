Embulk::JavaPlugin.register_formatter(
  "poi_excel", "org.embulk.formatter.poi_excel.PoiExcelFormatterPlugin",
  File.expand_path('../../../../classpath', __FILE__))
