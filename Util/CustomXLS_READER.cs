using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System;
using System.Linq;
using System.ComponentModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityQuickSheet;

public class CustomXLS_READER {
    private readonly IWorkbook workbook = null;
    private readonly ISheet sheet = null;
    private string filepath = string.Empty;
    public readonly List<ISheet> allSheets;
    /// <summary>
    /// Constructor.
    /// </summary>
    public CustomXLS_READER(string path, string sheetName = "")
    {
        try
        {
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                string extension = GetSuffix(path);
                if (extension == "xls")
                    workbook = new HSSFWorkbook(fileStream);
                else if (extension == "xlsx")
                {
#if UNITY_EDITOR_OSX
                        throw new Exception("xlsx is not supported on OSX.");
#else
                    workbook = new XSSFWorkbook(fileStream);
#endif
                }
                else
                {
                    throw new Exception("Wrong file.");
                }
                //NOTE: An empty sheetName can be available. Nothing to do with an empty sheetname.
                if (!string.IsNullOrEmpty(sheetName))
                    sheet = workbook.GetSheet(sheetName);
                this.filepath = path;
            }
        }
        catch (Exception e)
        {
            Debug.LogError(e.Message);
            Application.Quit();
        }
    }
    /// <summary>
    /// Determine whether the excel file is successfully read in or not.
    /// </summary>
    public bool IsValid()
    {
        if (this.workbook != null && this.sheet != null)
            return true;
        return false;
    }
    /// <summary>
    /// Retrieves file extension only from the given file path.
    /// </summary>
    static string GetSuffix(string path)
    {
        string ext = Path.GetExtension(path);
        string[] arg = ext.Split(new char[] { '.' });
        return arg[1];
    }
    public string GetHeaderColumnName(int cellnum)
    {
        ICell headerCell = sheet.GetRow(0).GetCell(cellnum);
        if (headerCell != null)
            return headerCell.StringCellValue;
        return string.Empty;
    }

    /// <summary>
    /// 
    /// Get informations from especific page and returns it as a string.
    /// Used for print data of sheet.
    ///
    /// <paramref name="page"/>
    /// Name of page to get data in the sheet.
    ///<paramref name="lang"/>
    /// Index of column with contains the specific values
    /// 
    /// NOTE: This function is used for traductions.
    /// 
    /// </summary>
    public string GetSheetString(string page){
        ISheet newSheet = workbook.GetSheet(page);
        string stringSheet = "";
        for(int row = 0; row < newSheet.PhysicalNumberOfRows;row++){
            if(newSheet.GetRow(row).PhysicalNumberOfCells == 0) 
                return stringSheet;
            stringSheet += row+" ";
            for(int col = 0; col < newSheet.GetRow(row).PhysicalNumberOfCells; col++){
                stringSheet += "| "+newSheet.GetRow(row).GetCell(col)+" ";
            }
            stringSheet += "|\n";
        }
        return stringSheet;        
    }    



    /// <summary>
    /// 
    /// Get informations of a specific page and column.
    ///
    ///<paramref name="lang"/>
    /// Index of column with contains the specific values
    /// 
    /// NOTE: This function is used for traductions.
    /// 
    /// </summary>
    public List<TraductionData> GetSheetInfo(int lang){
        List<TraductionData> data = new List<TraductionData>();
        for(int row = 1; row < sheet.PhysicalNumberOfRows;row++){
            if(sheet.GetRow(row).PhysicalNumberOfCells == 0) 
                return data;
            IRow Row = sheet.GetRow(row);
            data.Add(
                new TraductionData(
                    Row.GetCell(0).StringCellValue,
                    Row.GetCell(lang).StringCellValue
                )
            );
        }
        return data;
    }
    /// <summary>
    /// 
    /// Get informations of a specific page and column.
    ///
    /// <paramref name="page"/>
    /// Name of page to get data in the sheet.
    /// 
    /// NOTE: This function is used for traductions.
    /// 
    /// </summary>
    public List<TraductionData> GetSheetInfo(string page){
        ISheet newSheet = workbook.GetSheet(page);
        List<TraductionData> data = new List<TraductionData>();
        for(int row = 1; row < newSheet.PhysicalNumberOfRows;row++){
            if(newSheet.GetRow(row).PhysicalNumberOfCells == 0) 
                return data;
            IRow Row = newSheet.GetRow(row);
            data.Add(
                new TraductionData(
                    Row.GetCell(0).StringCellValue,
                    Row.GetCell(1).StringCellValue
                )
            );
        }
        return data;
    }

    /// <summary>
    /// 
    /// Get informations of a specific page and column.
    ///
    /// <paramref name="page"/>
    /// Name of page to get data in the sheet.
    ///<paramref name="lang"/>
    /// Index of column with contains the specific values
    /// 
    /// NOTE: This function is used for traductions.
    /// 
    /// </summary>
    public List<TraductionData> GetSheetInfo(string page, int lang){
        ISheet newSheet = workbook.GetSheet(page);
        List<TraductionData> data = new List<TraductionData>();
        for(int row = 1; row < newSheet.PhysicalNumberOfRows;row++){
            if(newSheet.GetRow(row).PhysicalNumberOfCells == 0) 
                return data;
            IRow Row = newSheet.GetRow(row);
            data.Add(
                new TraductionData(
                    Row.GetCell(0).StringCellValue,
                    Row.GetCell(lang).StringCellValue
                )
            );
        }
        return data;
    }      


    /// <summary>
    /// Deserialize all the cell of the given sheet.
    ///
    /// NOTE:
    ///     The first row of a sheet is header column which is not the actual value
    ///     so it skips when it deserializes.
    /// </summary>
    public List<T> Deserialize<T>(int start = 1)
    {
        var t = typeof(T);
        PropertyInfo[] p = t.GetProperties();
        var result = new List<T>();
        int current = 0;
        foreach (IRow row in sheet)
        {
            if (current < start)
            {
                current++; // skip header column.
                continue;
            }
            var item = (T)Activator.CreateInstance(t);
            for (var i = 0; i < p.Length; i++)
            {
                ICell cell = row.GetCell(i);
                var property = p[i];
                if (property.CanWrite)
                {
                    try
                    {
                        var value = ConvertFrom(cell, property.PropertyType);
                        property.SetValue(item, value, null);
                    }
                    catch (Exception e)
                    {
                        string pos = string.Format("Row[{0}], Cell[{1}]", (current).ToString(), GetHeaderColumnName(i));
                        Debug.LogError(string.Format("Excel File {0} Deserialize Exception: {1} at {2}", this.filepath, e.Message, pos));
                    }
                }
            }
            result.Add(item);
            current++;
        }
        return result;
    }
    /// <summary>
    /// Retrieves all sheet names.
    /// </summary>
    public string[] GetSheetNames()
    {
        List<string> sheetList = new List<string>();
        if (this.workbook != null)
        {
            int numSheets = this.workbook.NumberOfSheets;
            for (int i = 0; i < numSheets; i++)
            {
                sheetList.Add(this.workbook.GetSheetName(i));
            }
        }
        else
            Debug.LogError("Workbook is null. Did you forget to import excel file first?");
        return (sheetList.Count > 0) ? sheetList.ToArray() : null;
    }
    public List<ISheet> GetAllSheets()
    {
        List<ISheet> s = new List<ISheet>();
        if (this.workbook != null)
        {
            int numSheets = this.workbook.NumberOfSheets;
            for (int i = 0; i < numSheets; i++)
            {
                s.Add(workbook.GetSheet(workbook.GetSheetName(i)));
            }
        }
        return s;
    }
    /// <summary>
    /// Retrieves all first columns(aka. header column) which are needed to determine each type of a cell.
    /// </summary>
    public string[] GetTitle(int start, ref string error)
    {
        List<string> result = new List<string>();
        IRow title = sheet.GetRow(start);
        if (title != null)
        {
            for (int i = 0; i < title.LastCellNum; i++)
            {
                string value = title.GetCell(i).StringCellValue;
                if (string.IsNullOrEmpty(value))
                {
                    error = string.Format(@"Empty column is found at {0}.", i);
                    return null;
                }
                else
                {
                    // column header is not an empty string, we check its validation later.
                    result.Add(value);
                }
            }
            return result.ToArray();
        }
        error = string.Format(@"Empty row at {0}", start);
        return null;
    }
    /// <summary>
    /// Retrieves all first columns(aka. header column) which are needed to determine each type of a cell.
    /// </summary>
    public string[] GetTitle(int start)
    {
        List<string> result = new List<string>();
        IRow title = sheet.GetRow(start);
        if (title != null)
        {
            for (int i = 0; i < title.LastCellNum; i++)
            {
                string value = title.GetCell(i).StringCellValue;
                if (string.IsNullOrEmpty(value))
                {
                    result.Add("");
                }
                else
                {
                    // column header is not an empty string, we check its validation later.
                    result.Add(value);
                }
            }
            return result.ToArray();
        }
        return null;
    }
    /// <summary>
    /// Convert type of cell value to its predefined type which is specified in the sheet's ScriptMachine setting file.
    /// </summary>
    protected object ConvertFrom(ICell cell, Type t)
    {
        object value = null;
        if (t == typeof(float) || t == typeof(double) || t == typeof(short) || t == typeof(int) || t == typeof(long))
        {
            if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
            {
                value = cell.NumericCellValue;
            }
            else if (cell.CellType == NPOI.SS.UserModel.CellType.String)
            {
                //Get correct numeric value even the cell is string type but defined with a numeric type in a data class.
                if (t == typeof(float))
                    value = Convert.ToSingle(cell.StringCellValue);
                if (t == typeof(double))
                    value = Convert.ToDouble(cell.StringCellValue);
                if (t == typeof(short))
                    value = Convert.ToInt16(cell.StringCellValue);
                if (t == typeof(int))
                    value = Convert.ToInt32(cell.StringCellValue);
                if (t == typeof(long))
                    value = Convert.ToInt64(cell.StringCellValue);
            }
            else if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
            {
                // Get value even if cell is a formula
                if (t == typeof(float))
                    value = Convert.ToSingle(cell.NumericCellValue);
                if (t == typeof(double))
                    value = Convert.ToDouble(cell.NumericCellValue);
                if (t == typeof(short))
                    value = Convert.ToInt16(cell.NumericCellValue);
                if (t == typeof(int))
                    value = Convert.ToInt32(cell.NumericCellValue);
                if (t == typeof(long))
                    value = Convert.ToInt64(cell.NumericCellValue);
            }
        }
        else if (t == typeof(string) || t.IsArray)
        {
            // HACK: handles the case that a cell contains numeric value
            //       but a member field in a data class is defined as string type.
            //       e.g. string s = "123"
            if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                value = cell.NumericCellValue;
            else
                value = cell.StringCellValue;
        }
        else if (t == typeof(bool))
            value = cell.BooleanCellValue;
        if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
        {
            var nc = new NullableConverter(t);
            return nc.ConvertFrom(value);
        }
        if (t.IsEnum)
        {
            // for enum type, first get value by string then convert it to enum.
            value = cell.StringCellValue;
            return Enum.Parse(t, value.ToString(), true);
        }
        else if (t.IsArray)
        {
            if (t.GetElementType() == typeof(float))
                return ConvertExt.ToSingleArray((string)value);
            if (t.GetElementType() == typeof(double))
                return ConvertExt.ToDoubleArray((string)value);
            if (t.GetElementType() == typeof(short))
                return ConvertExt.ToInt16Array((string)value);
            if (t.GetElementType() == typeof(int))
                return ConvertExt.ToInt32Array((string)value);
            if (t.GetElementType() == typeof(long))
                return ConvertExt.ToInt64Array((string)value);
            if (t.GetElementType() == typeof(string))
                return ConvertExt.ToStringArray((string)value);
        }
        // for all other types, convert its corresponding type.
        return Convert.ChangeType(value, t);
    }
}
