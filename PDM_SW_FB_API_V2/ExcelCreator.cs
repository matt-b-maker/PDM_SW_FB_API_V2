using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PDM_SW_FB_API_V2
{
    class ExcelCreator
    {
        public ExcelCreator (string prodNum, string path, List<MaterialItem> materials, List<PurchasedItem> purchasedItems)
        {
            this.ProdNum = prodNum;
            this.Path = path;
            this.Materials = materials;
            this.PurchasedItems = purchasedItems;
        }

        public ExcelCreator() { }

        public string ProdNum { get; set; }
        public string Path { get; set; }
        public FileInfo File { get; set; }
        public List<MaterialItem> Materials { get; set; }
        public List<PurchasedItem> PurchasedItems { get; set; }

        public FileInfo CreateExcelFile(string path, string prodNum)
        {
            var excelFile = new FileInfo(fileName: @"" + path + $"\\PROD-{prodNum} FB BOM.xlsx");
            return excelFile;
        }

        public async Task SaveExcelFile(List<MaterialItem> materialItems, List<PurchasedItem> purchasedItems, FileInfo file)
        {
            DeleteIfExists(file);

            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add(Name: "FB BOM");

                for (int i = 0; i < materialItems.Count; i++)
                {
                    ws.Cells[Address: $"A{i + 1}"].Value = materialItems[i].PartNo;
                    ws.Cells[Address: $"B{i + 1}"].Value = materialItems[i].Quantity;
                }
                //Have to carry on in the spreadsheet using the number of material items as a modifier
                for (int i = materialItems.Count; i < (materialItems.Count + purchasedItems.Count); i++)
                {
                    ws.Cells[Address: $"A{i + 1}"].Value = purchasedItems[i - materialItems.Count].PartNo;
                    ws.Cells[Address: $"B{i + 1}"].Value = purchasedItems[i - materialItems.Count].Quantity;
                }

                ws.Column(1).AutoFit();
                ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(2).AutoFit();
                ws.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                await package.SaveAsync();
            }
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        public void OpenExcelDoc(FileInfo file)
        {
            //Process.Start(@"cmd.exe", @"/c" + file.FullName);

            new Process { StartInfo = new ProcessStartInfo(@"" + file.FullName) { UseShellExecute = true } }.Start();
        }

        
    }
}
