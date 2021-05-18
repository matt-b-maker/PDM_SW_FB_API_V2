using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using EPDM.Interop.epdm;

namespace PDM_SW_FB_API_V2
{
    class Program
    {
        static void Main()
        {
            #region Set up variables

            //For region 1
            bool loggedIn = false;
            bool exceptionEncountered = false;

            IEdmFile5 File;
            IEdmReference5 Reference;
            IEdmReference10 @ref;

            string prodNum;

            IEdmVault21 CurrentVault = new EdmVault5() as IEdmVault21;
            IEdmSearch9 _search;
            IEdmSearchResult5 _searchResult;

            //For region 2
            List<PurchasedItem> purchasedItems = new List<PurchasedItem>();

            //For region 3
            string fullPath;
            string cncPath;

            #endregion

            #region 1. Search PDM for PROD # based off of user input

            do
            {
                Console.WriteLine("Enter the four numbers of the PROD #: ");
                prodNum = Console.ReadLine();

                try
                {
                    CurrentVault.LoginAuto("CreativeWorks", 0);
                    loggedIn = true;
                }
                catch
                {
                    Console.WriteLine("You need to be logged into the PDM, genius. \n\n");
                }
            } while (loggedIn == false);

            _search = (IEdmSearch9)CurrentVault.CreateSearch2();
            _search.FindFiles = true;

            _search.Clear();
            _search.StartFolderID = CurrentVault.GetFolderFromPath("C:\\CreativeWorks").ID;

            //_search.AddVariable("DocumentNumber", prodNum);
            _search.FileName = "PROD-" + prodNum + ".sldasm";

            _search.GetFirstResult();

            if (exceptionEncountered)
            {
                _searchResult = null;
            }
            else
            {
                _searchResult = _search.GetFirstResult();
            }

            if (_searchResult == null)
            {
                Console.WriteLine("Didn't find anything for that");
                Console.ReadKey();
                return;
            }
            else
            {
                fullPath = _searchResult.Path;
            }

            #endregion

            #region 2. Find hardware and populate Purchased Item object list

            File = CurrentVault.GetFileFromPath(_searchResult.Path, out IEdmFolder5 ParentFolder);

            Reference = (IEdmReference10)File.GetReferenceTree(ParentFolder.ID);

            IEdmPos5 pos = Reference.GetFirstChildPosition("Get Some", true, true, 0);
            //Console.WriteLine(pos.ToString());x 

            while (!pos.IsNull)
            {
                @ref = (IEdmReference10)Reference.GetNextChild(pos);
                if (!@ref.Name.Contains("PROD"))
                {
                    purchasedItems.Add(new PurchasedItem(GetPartNo(@ref.Name), @ref.RefCount));
                }
            }

            if (purchasedItems.Count > 0)
            {
                foreach (var item in purchasedItems)
                {
                    Console.WriteLine(item.PartNo + " " + item.Quantity);
                }
            }
            else
            {
                Console.WriteLine("There wasn't any hardware that could be found for that PROD");
                Console.WriteLine("Press any key to exit");
                Console.ReadKey();
                return;
            }

            #endregion

            #region 3. Use PROD# to find CNC folder and count sheet materials inside

            cncPath = GetCncPath(fullPath);

            #endregion

            Console.ReadKey();
        }

        private static string GetPartNo(string name)
        {
            string partNoPattern1 = @".SLDPRT";
            string partNoPattern2 = @"[a-zA-Z0-9]+$";
            string partNo = Regex.Replace(name, partNoPattern1, "");
            partNo = Regex.Match(partNo, partNoPattern2).ToString();
            return partNo;
        }

        private static string GetCncPath(string fullPath)
        {
            string cutPattern = @"1-Design Reference.+";
            string cncPath = Regex.Replace(fullPath, cutPattern, "");
            cncPath += "2-CNC";
            return cncPath;
        }
    }
}
