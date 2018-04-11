using System;
using System.Collections.Generic;
using System.Collections;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace ExcelToExcel
{
    static class Tools
    {
        public static void SelectExcelFile(OpenFileDialog file)
        {

        }

        public static void ErrorBox(Exception x)
        {
            MessageBox.Show(x.ToString(), "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void CloseExcel(Excel.Application app,
                                      Excel.Workbook wb,
                                      Label l,
                                      Button bLoad)
        {
            DialogResult result = MessageBox.Show("Voulez-vous sauvez le fichier?", 
                "Sauver Fichier", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);

            try
            {
                if (result == DialogResult.No)
                {
                    wb.Close(false);

                    app.Quit();
                    wb = null;

                    app = null;
                    l.Text = "";

                    bLoad.Enabled = true;
                }
                else if (result == DialogResult.Yes)
                {
                    wb.Close(true);

                    app.Quit();
                    wb = null;

                    app = null;
                    l.Text = "";

                    bLoad.Enabled = true;
                }
                else
                { }
            }
            catch
            {
                MessageBox.Show("Votre fichier Excel est déjà fermé.", 
                    "Information", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Information);
            }
        }

        public static void RemoveCheckedNodes(TreeNodeCollection nodes)
        {
            for (int i = nodes.Count - 1; i >= 0; --i)
            {
                if (nodes[i].Checked)
                {
                    nodes.RemoveAt(i);
                }
                else
                {
                    RemoveCheckedNodes(nodes[i].Nodes);
                }
            }
        }


        public static Hierarchie Define(TreeNode mainNode, int n, int i, ListBox listBox, Excel.Worksheet xlWs)
        {
            int j = 0;
            List<Hierarchie> ssHierarchie = new List<Hierarchie>();

            foreach (TreeNode subNode in mainNode.Nodes)
            {
                j++;
                ssHierarchie.Add(Define(subNode, n + 1, j, listBox, xlWs));
            }

            if (j == 0)
            {
                Local local = new Local()
                {
                    Nom = mainNode.Text,
                    IdRelationnel = i,
                    Niveau = n,
                    Cell = (Excel.Range)mainNode.Tag,
                };
                local.Properties = GiveProperties(listBox, local, xlWs);

                mainNode.ToolTipText = local.ToString();
                mainNode.Tag = local;

                return local;
            }
            else
            {
                Hierarchie hierarchie = new Hierarchie(mainNode.Text, i, n, ssHierarchie);

                mainNode.ToolTipText = hierarchie.ToString();
                mainNode.Tag = hierarchie;

                return hierarchie;
            }
        }

        public static List<Tuple<string, string>> GiveProperties(ListBox myListBox, Local local, Excel.Worksheet xlWs)
        {
            List<Tuple<string, string>> listProperties = new List<Tuple<string, string>>();
            foreach (Property property in myListBox.Items)
            {
                Property p = new Property()
                {
                    CellNom = property.CellNom,
                    Nom = property.Nom,
                };

                int propertyColumn = p.CellNom.Column;
                int localRow = local.Cell.Row;
                p.CellValue = xlWs.Cells[RowIndex: localRow, ColumnIndex: propertyColumn];

                try
                {
                    p.Value = p.CellValue.Value2?.ToString();
                }
                catch (Exception x)
                {
                    ErrorBox(x);
                }
                listProperties.Add(Tuple.Create(p.Nom, p.Value));               
            }
            return listProperties;
        }

        public static DataTable CreateDataTableHeaders(List<Property> listProperties)
        {
            List<string> entetes = new List<string>
            {
                "Id_Entité",
                "Entité",

                "Id_Secteur",
                "Secteur",

                "Id_Sous-Secteur",
                "Sous-Secteur",

                "Id_Zone",
                "Zone",

                "Id_Sous-Zone",
                "Sous-Zone",

                "Id_Groupe",
                "Groupe",

                "Id_Sous-Groupe",
                "Sous-Groupe",

                "Id_Local_Unique",
                "Id_Local_Relationnel",
                "Local"
            };

            DataTable myData = new DataTable();
            DataColumn dataColumn;

            //Create a header for each hierarchy level
            for (int i = 0; i < entetes.Count; i++)
            {
                dataColumn = new DataColumn()
                {
                    ColumnName = entetes[i],
                };
                myData.Columns.Add(dataColumn);
            }

            //Create a header for each property 
            for (int i = 0; i < listProperties.Count; i++)
            {
                dataColumn = new DataColumn()
                {
                    ColumnName = listProperties[i].Nom,
                };
                myData.Columns.Add(dataColumn);
            }

            return myData;
        }

        static int nbLocal = 1;
        public static void PopulateDataTable(DataTable myData, Hierarchie hierarchie, ArrayList listHierarchie)
        {
            listHierarchie = new ArrayList(listHierarchie);

            DataRow dataRow;
        
            if (hierarchie is Local local)
            {
                //Check for "Nb" and duplicate if >1
                int nbRepetition = 1;
                for (int i = 0; i < local.Properties.Count; i++)
                {
                    if (local.Properties[i].Item1 == "Nb" && local.Properties[i].Item2 != "")
                    {
                        if (Convert.ToInt16(local.Properties[i].Item2) > 1)
                        {
                            nbRepetition = Convert.ToInt16(local.Properties[i].Item2);
                            local.Properties[i] = new Tuple<string, string>(local.Properties[i].Item1, "1");
                        }
                    }
                }
                
                //Repeat depending on number
                for (int repet = 1; repet <= nbRepetition; repet++)
                {
                    dataRow = myData.NewRow();

                    dataRow["Id_Local_Relationnel"] = local.IdRelationnel;
                    dataRow["Local"] = local.Nom;
                    dataRow["Id_Local_Unique"] = nbLocal;

                    nbLocal++;

                    //Add Id_Entité -> Sous-groupe
                    for (int i = 0; i < listHierarchie.Count; i++)
                    {
                        dataRow[i] = listHierarchie[i]; 
                    }

                    //Add local properties to datatable
                    foreach(Tuple<string, string> property in local.Properties)
                    {
                        if(!myData.Columns.Contains(property.Item1))
                        {
                            myData.Columns.Add(property.Item1);
                        }
                        dataRow[property.Item1] = property.Item2;
                    }

                    myData.Rows.Add(dataRow);
                }
            }
            else
            {
                listHierarchie.Add((hierarchie.IdEntite == 0) ? hierarchie.Id : hierarchie.IdEntite);
                listHierarchie.Add(hierarchie.Nom);
                foreach (Hierarchie ssHierarchie in hierarchie.List)
                {
                    PopulateDataTable(myData, ssHierarchie, listHierarchie);
                }
            }
        }

    }
}
