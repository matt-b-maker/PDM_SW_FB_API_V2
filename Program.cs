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
            List<MaterialItem> materialItems = new List<MaterialItem>();
            string temp = null; 

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
                Console.WriteLine();
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
                Console.WriteLine();
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

            _search.Clear();
            _search.StartFolderID = CurrentVault.GetFolderFromPath(cncPath).ID;

            _search.FindFiles = true;

            _search.FileName = ".sbp";

            _searchResult = _search.GetFirstResult();

            if (_searchResult != null && !_searchResult.Path.Contains("Parts"))
            {
                temp = GetMaterialName(_searchResult.Name);
                materialItems.Add(new MaterialItem(temp, 1, GetMaterialPartNo(temp)));
            }

            while (_searchResult != null)
            {
                _searchResult = _search.GetNextResult();

                if (_searchResult != null && _searchResult.Path.Contains("Programs") && !_searchResult.Path.Contains("Parts") && !IsMultiPart(_searchResult.Name))
                {
                    if (GetMaterialName(_searchResult.Name) != temp)
                    {
                        temp = GetMaterialName(_searchResult.Name);
                        materialItems.Add(new MaterialItem(temp, 1, GetMaterialPartNo(temp)));
                    }
                    else
                    {
                        materialItems[materialItems.Count - 1].Quantity++;
                    }
                }
                else
                {
                    continue;
                }
            }

            if (materialItems.Count > 0)
            {
                Console.Write("Material Name                 ");
                Console.Write("Quantity                      ");
                Console.Write("Part Number in Fishbowl       " + "\n");

                DrawLine(13, 30);
                DrawLine(8, 30);
                DrawLine(23, 30);

                Console.WriteLine();

                foreach (var item in materialItems)
                {
                    WriteMaterialInfo(item);
                }
            }

            #endregion

            Console.ReadKey();
        }

        private static string GetMaterialPartNo(string temp)
        {
            switch (temp)
            {
                case "Celtec-black 50":
                    return "12760";
                case "1-MDX 50":
                    return "0214003";
                case "2-MDX 75":
                    return "222003";
                case "3-Cheap Plywood 18mm (.7) ":
                    return "518447";
                case "NO MARGIN - Cheap Plywood 18mm (.7)":
                    return "518447";
                case "3mm Laminate":
                    return "Find out which laminate";
                case "4- Birch Plywood 18mm (.7)":
                    return "0622133";
                case "ABS-black 125":
                    return "173144";
                case "ABS-white 125":
                    return "WHITEABS";
                case "Acrylic-clear (Impact Resistant) 25":
                    return "1/4\" Super_Impact Resistant acrylic";
                case "Acrylic-clear 25":
                    return "220958";
                case "Acrylic-clear 375":
                    return "16681";
                case "Acrylic-clear 50":
                    return "244373";
                case "Acrylic-mirrored 125":
                    return "18266";
                case "Acrylic-orange 125":
                    return "ACRY21190";
                case "Acrylic-orange 25":
                    return "ACS-00220ORG000001";
                case "Bendaply 375 (Hamburg)":
                    return "6041";
                case "Plywood 375 bendaply (Grain)":
                    return "6041";
                case "Bendaply 375 (Hotdog)":
                    return "6040";
                case "Celtec-black 75":
                    return "15715";
                case "Celtec-white 25":
                    return "13428";
                case "Celtec-white 50":
                    return "208138";
                case "Celtec-white 75":
                    return "208137";
                case "ColorCore Blue/White/Blue 75":
                    return "36916";
                case "Dibond 125":
                    return "184514";
                case "HDPE-white 0625":
                    return "18096";
                case "HDPE-white 125":
                    return "111717";
                case "HDPE-white 25":
                    return "111723";
                case "HDPE-white 75":
                    return "39175";
                case "MDF 25":
                    return "3908531";
                case "MDF 25 no Margin":
                    return "3908531";
                case "MDF 25 Stencil Board":
                    return "3908531";
                case "MDF 50":
                    return "MDF1248";
                case "MDF 75":
                    return "MDF 3/4";
                case "MDF 75 for 10' Board":
                    return "MDF 3/4";
                case "MDX 75 NO MARGIN":
                    return "MDF 3/4";
                case "Melamine-black 18mm":
                    return "922133";
                case "Particle 6875":
                    return "8022539";
                case "PETG 0625":
                    return "18928";
                case "Plywood 6875":
                    return "18MMBirchAetna";
                case "Plywood 6875 (Grain)":
                    return "18MMBirchAetna";
                case "Seaboard-black 50":
                    return "33770";
                case "Seaboard-black 75":
                    return "25979";
                default:
                    return "N/A";
            }
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

        private static bool IsMultiPart(string name)
        {
            string pattern = @"[0-9]\.sbp";
            string checkName = Regex.Match(name, pattern).ToString();

            if (checkName != null)
            {
                if (checkName.StartsWith('1'))
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }

            return true;
        }

        private static string GetMaterialName(string name)
        {
            string frontEndPattern = @"^[0-9]+_";
            string backEndPattern = @"[0-9].sbp";

            string materialName = Regex.Replace(name, frontEndPattern, "");
            materialName = Regex.Replace(materialName, backEndPattern, "");

            return materialName;
        }

        private static void DrawLine(int num, int max)
        {
            for (int i = 0; i < num; i++)
            {
                Console.Write("-");
            }
            for (int j = 0; j < max - num; j++)
            {
                Console.Write(" ");
            }
        }

        private static void WriteMaterialInfo(MaterialItem item)
        {
            Console.Write(item.MaterialName);
            for (int i = 0; i < 30 - item.MaterialName.Length; i++)
            {
                Console.Write(" ");
            }
            Console.Write(item.Quantity);
            for (int i = 0; i < 30 - item.Quantity.ToString().Length; i++)
            {
                Console.Write(" ");
            }
            Console.Write(item.PartNo + "\n");
        }
    }
}
