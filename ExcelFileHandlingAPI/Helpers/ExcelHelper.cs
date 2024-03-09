using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelFileHandlingAPI.Helpers
{
    public static class ExcelHelper
    {
        public static List<T> Extract<T>(string filePath) where T : new()
        {
            XSSFWorkbook workbook;
            using(var file = new FileStream(filePath,FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }

            var sheet = workbook.GetSheetAt(0);

            var rowHeader = sheet.GetRow(0);
            var headerList = new Dictionary<string,int>();
            foreach(var cell in rowHeader.Cells)
            {
                var collName = cell.StringCellValue;
                headerList.Add(collName, cell.ColumnIndex);
            }

            var listObject = new List<T>();

            for(int i = 1;i<sheet.LastRowNum;i++)
            {
                var row = sheet.GetRow(i);
               
                if(row == null) break;

                var obj = new T();

                foreach(var property in typeof(T).GetProperties())
                {
                    var colIndex = headerList[property.Name];
                    var cell = row.GetCell(colIndex);

                    if(cell == null)
                    {
                        property.SetValue(obj, null);
                    }
                    else if (property.PropertyType == typeof(string))
                    {
                        cell.SetCellType(CellType.String);
                        property.SetValue(obj, cell.StringCellValue);
                    }
                    else if(property.PropertyType == typeof (int))
                    {
                        cell.SetCellType(CellType.Numeric);
                        property.SetValue(obj, Convert.ToInt32(cell.NumericCellValue));
                    }
                    else if(property.PropertyType == typeof(decimal))
                    {
                        cell.SetCellType(CellType.Numeric);
                        property.SetValue(obj, Convert.ToDecimal(cell.NumericCellValue));
                    }
                    else if(property.PropertyType == typeof(double))
                    {
                        cell.SetCellType(CellType.Numeric);
                        property.SetValue(obj, Convert.ToDouble(cell.NumericCellValue));
                    }

                }

                listObject.Add(obj);
            }
                
            return listObject;
        }
    }
}
