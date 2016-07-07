using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;
using NPOI.SS.Util;

namespace TestProject
{
    class WorkbookUtil
    {
        Dictionary<string, string> props = new Dictionary<string, string>();
 

        public void abc(String FilePath)
        {

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(FilePath), FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

         
            ISheet sheet = hssfwb.GetSheet("Sheet1");
         

           
           
            for (int row = 0; row <= sheet.LastRowNum+1; row++)
            {
                if (sheet.GetRow(row) != null) 
                {
                   /* foreach (var cell in sheet.GetRow(row).Cells)
                    {
                        int col=cell.ColumnIndex;
                    }
                    */
                  

		/*
		//if Name range is used in the formula , then Regex can be used to extract the name and formula for name be determined as below
			 var regex_Collection = Regex.Matches(sheet.GetRow(row).GetCell(col).CellFormula, "_\\w+");
                    String  formula = sheet.GetRow(row).GetCell(Col).CellFormula;
                    XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(hssfwb);
                   
			 foreach (Match item in regex_Collection)
                    {
                     
                        //var regex_Collection2=Regex.Matches(hssfwb.GetName(item.Value).RefersToFormula, "(\\d)[^\\d]*$");
                 

                        AreaReference area = new AreaReference(Regex.Replace(hssfwb.GetName(item.Value).RefersToFormula, "(\\d)[^\\d]*$", (row+1).ToString()));

                        formula= formula.Replace(item.Value, area.ToString().Split(' ')[1].Replace("[","").Replace("]",""));
                      
                    }
                    sheet.GetRow(row).GetCell(Col).CellFormula = formula;

                      CellValue abc = evaluator.Evaluate(sheet.GetRow(row).GetCell(col));
                                      
                }
            }

            

        }
*/
//to read the excel data using oledb connection

        public DataTable GetDataTableFromExcelFile(string fullFileName, string sheetName)
        {
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            props["Data Source"] = fullFileName;
            props["Extended Properties"] = "Excel 12.0";
            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }
            string properties = sb.ToString();

            OleDbConnection objConnection = new OleDbConnection();
            objConnection = new OleDbConnection(properties);

            using (OleDbConnection conn = new OleDbConnection(properties))
            {
                conn.Open();
                //Get All Sheets Name
                DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });

                //Get the First Sheet Name
                string firstSheetName = sheetsName.Rows[0][2].ToString();

                //Query String 
                string sql = string.Format("SELECT * FROM [{0}]", firstSheetName);
                OleDbDataAdapter ada = new OleDbDataAdapter(sql, properties);
                DataSet set = new DataSet();
                ada.Fill(set);
            
               
                return set.Tables[0];
            }

        }
    }
}



