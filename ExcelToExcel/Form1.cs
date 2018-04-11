using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;

namespace ExcelToExcel
{
    public partial class MainForm : Form
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;

        private List<Hierarchie> listHierarchie = new List<Hierarchie>();

        private List<Property> listProperties = new List<Property>();
        private bool propertyListStateOnCreate = false;

        DataTable myData;
        DataTable ficheEspaceData;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult closeApp = MessageBox.Show("Voulez-vous vraiment fermer l'application?", 
                "Confirmation", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);

            if (closeApp == DialogResult.Yes)
            {
                if (xlApp == null && xlWorkbook == null)
                { return; }
                DialogResult saveExcel = MessageBox.Show("Voulez-vous sauver le fichier Excel?",
                    "Sauver",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                bool save = (saveExcel == DialogResult.Yes) ? true : false;
                try
                {
                    xlWorkbook.Close(save);
                }
                catch
                {
                    MessageBox.Show("Impossible de sauvegarder: vous avez déjà fermé le fichier Excel.",
                        "Erreur",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                #region:CloseExcelClean
                xlApp.Quit();

                try
                {
                    while (Marshal.ReleaseComObject(xlWorkbook) > 0)
                    {
                    }
                }
                catch
                { }
                finally
                {
                    xlWorkbook = null;
                }

                try
                {
                    while (Marshal.ReleaseComObject(xlApp) > 0)
                    {
                    }
                }
                catch
                { }
                finally
                {
                    xlApp = null;
                }

                GC.Collect();
                #endregion:CloseExcelClean
            }
            else
            {
                e.Cancel = (closeApp == DialogResult.No);
            }  
        }

        #region:TAB-ExcelFile
        private void BT_Load_Click(object sender, EventArgs e)
        {
            #region:SelectExcelFile
            OF_xls.Title = "Sélectionner fichier Excel";
            OF_xls.Filter = "Microsoft Excel Worksheet 2007 (*.xlsx)|*.xlsx| " +
                            "Microsoft Excel Worksheet 2003 (*.xls)|*.xls";
            OF_xls.RestoreDirectory = true;
            OF_xls.ShowDialog();
            #endregion:SelectExcelFile

            #region:OpenExcelFile
            xlApp = new Excel.Application()
            {
                Visible = true
            };
            xlWorkbook = xlApp.Workbooks.Open(L_Path.Text, 3, 1);
            MessageBox.Show("Ouverture du fichier \"" + xlWorkbook.Name + "\" réussie",
                            "Success",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
            #endregion:OpenExcelFile

            #region:AddExcelWorksheetsToTreeView
            foreach(Excel.Worksheet wSheet in xlWorkbook.Worksheets)
            {
                TreeNode tn = new TreeNode()
                {
                    Text = wSheet.Name,
                    ToolTipText = "0",
                };
                myTreeView.Nodes.Add(tn);
            }
            #endregion:AddExcelWorksheetsToTreeView
        }

        private void OF_xls_FileOk(object sender, CancelEventArgs e)
        {
            L_Path.Text = OF_xls.FileName;
        }
        #endregion:TAB-ExcelFile

        #region:TAB-Hierarchy
        private void BT_Create_Click(object sender, EventArgs e)
        {
            if (myTreeView.SelectedNode != null)
            {
                //Add all nodes to this entity (selected node is changed)
                TreeNode parentNode = myTreeView.SelectedNode;

                #region:CreateNodes
                Excel.Range selection = xlApp.Selection;

                List<Cell_Style> styleList = new List<Cell_Style>();

                for (int i = 1; i <= selection.Rows.Count; i++)
                {
                    TreeNode treeNode;

                    Excel.Range cell = selection.Cells[i];
                    if (cell.Value2 == null)
                    {
                        continue;
                    }

                    Cell_Style thisCell = new Cell_Style(cell);                    

                    if (!styleList.Contains(thisCell))
                    {
                        styleList.Add(thisCell);
                    }
                    else if (styleList[styleList.Count - 1].Equals(thisCell))
                    {
                        myTreeView.SelectedNode = myTreeView.SelectedNode.Parent;
                    }
                    else
                    {
                        int indexPreviousCellofSameStyle = styleList.IndexOf(thisCell);

                        //Select TN parent
                        for (int j = 1; j <= (styleList.Count - indexPreviousCellofSameStyle); j++)
                        {
                            myTreeView.SelectedNode = myTreeView.SelectedNode.Parent;
                        }

                        styleList.RemoveRange(indexPreviousCellofSameStyle + 1, styleList.Count - indexPreviousCellofSameStyle - 1);
                    }

                    //Add new TN to parent - TN object corresponds to excel cell
                    treeNode = new TreeNode()
                    {
                        Text = selection.Cells[i].Value2,
                        Tag = selection.Cells[i],
                    };
                    myTreeView.SelectedNode.Nodes.Add(treeNode);
                    treeNode.ToolTipText = treeNode.Level.ToString();

                    //selected TN => created TN
                    myTreeView.SelectedNode = treeNode;
                }
                #endregion:CreateNodes

                //Add nodes to a list of all hierarchies
                listHierarchie.Add(Tools.Define(parentNode, 0, 1, LB_Entetes, xlWorkbook.ActiveSheet));
                
                if (LB_Entetes.Items.Count > 0)
                {
                    //Create a non-modifiable & unique list of properties -> used for creating dataset heasers
                    if (propertyListStateOnCreate == false)
                    {
                        foreach (Property property in LB_Entetes.Items)
                        {
                            listProperties.Add(property);
                        }
                        propertyListStateOnCreate = true;
                    }

                    LB_Entetes.Items.Clear();
                }
            }
            else
            {
                MessageBox.Show("Sélectionnez une entité dans laquelle ajouter les données", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }          
        }

        private void BT_Delete_Node_Click(object sender, EventArgs e)
        {
            //Delete checked nodes and all children nodes
            Tools.RemoveCheckedNodes(myTreeView.Nodes);
        }

        private void BT_collapse_Click(object sender, EventArgs e)
        {
            myTreeView.CollapseAll();
        }

        private void BT_expand_Click(object sender, EventArgs e)
        {
            myTreeView.ExpandAll();
        }

        private void BT_Add_Property_Click(object sender, EventArgs e)
        {
            if (xlApp.Selection.Cells != null)
            {
                //Add properties to listBox
                Property p;
                foreach (Excel.Range cell in xlApp.Selection.Cells)
                {
                    p = new Property();

                    if (cell.Value2 != null)
                    {
                        p.CellNom = cell;
                        p.Nom = cell.Value2.ToString();

                        LB_Entetes.Items.Add(p);
                    }
                }
            }
            else
            {
                MessageBox.Show("Vous n'avez pas sélectionné les entêtes à ajouter.", "Alerte", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            
        }

        private void BT_Delete_Property_Click(object sender, EventArgs e)
        {
            if (LB_Entetes.SelectedItems != null)
            {
                for (int x = LB_Entetes.SelectedIndices.Count - 1; x >= 0; x--)
                {
                    int idx = LB_Entetes.SelectedIndices[x];
                    LB_Entetes.Items.RemoveAt(idx);
                }
            }
            else
            {
                MessageBox.Show("Vous n'avez pas sélectionné d'entêtes à supprimer", "Alerte", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void RB_FE_False_CheckedChanged(object sender, EventArgs e)
        {
            TB_FE_Name.Enabled = false;
            TB_FE_UL.Enabled = false;
            TB_FE_BR.Enabled = false;

            BT_FE_Confirm.Enabled = false;
        }

        private void RB_FE_True_CheckedChanged(object sender, EventArgs e)
        {
            TB_FE_Name.Enabled = true;
            TB_FE_UL.Enabled = true;
            TB_FE_BR.Enabled = true;

            BT_FE_Confirm.Enabled = true;
        }

        private void BT_FE_Confirm_Click(object sender, EventArgs e)
        {
            //Creates a dataTable corresponding to fiche espace
            if (TB_FE_Name.Text != "" && TB_FE_Name.Text != null)
            {
                if (TB_FE_UL.Text != "" && TB_FE_UL != null)
                {
                    if (TB_FE_BR.Text != "" && TB_FE_UL != null)
                    {
                        #region:CheckExistanceOfSheet
                        Excel.Worksheet wsFicheEspace = new Excel.Worksheet();
                        try
                        {
                            wsFicheEspace = xlWorkbook.Worksheets.Item[TB_FE_Name.Text];
                        }
                        catch
                        {
                            string errorMessage = "La feuille de calcul " + TB_FE_Name.Text + " n'existe pas";
                            MessageBox.Show(errorMessage, 
                                "Erreur", 
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return;
                        }
                        #endregion:CheckExistanceOfSheet

                        #region:CheckCorrectnessOfRange
                        Excel.Range rangeFicheEspace;
                        try
                        {
                            rangeFicheEspace = wsFicheEspace.Range[TB_FE_UL.Text, TB_FE_BR.Text];
                        }
                        catch
                        {
                            string errorMessage = "La portée des données est incorrecte [" + TB_FE_UL.Text + ";" + "]";
                            MessageBox.Show(errorMessage, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        #endregion:CheckCorrectnessOfRange

                        #region:CreateDataTable
                        ficheEspaceData = new DataTable();
                        for (int i = 2; i <= rangeFicheEspace.Rows.Count; i++)
                        {
                            DataRow ficheEspaceRow = ficheEspaceData.NewRow();
                            string columnName;
                            string columnValue;

                            for (int j = 1; j <= rangeFicheEspace.Columns.Count; j++)
                            {
                                columnName = rangeFicheEspace[RowIndex: 1, ColumnIndex: j].Value2?.ToString();
                                columnValue = rangeFicheEspace[RowIndex: i, ColumnIndex: j].Value2?.ToString();

                                if (!ficheEspaceData.Columns.Contains(columnName))
                                {
                                    ficheEspaceData.Columns.Add(columnName);
                                }
                                ficheEspaceRow[columnName] = columnValue;
                            }
                            ficheEspaceData.Rows.Add(ficheEspaceRow);
                        }
                        MessageBox.Show("Importation de la fiche espace effectuée.",
                            "Information",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        #endregion:CreateDataTable
                    }
                    else
                    {
                        MessageBox.Show("Veuillez remplir le champ correspondant à la cellule en bas à droite de la fiche espace",
                            "Erreur",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Veuillez remplir le champ correspondant à la cellule en haut à gauche de la fiche espace",
                        "Erreur",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Veuillez remplir le champ correspondant au nom de la feuille de calcul de la fiche espace",
                    "Erreur",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        #endregion:TAB-Hierarchy

        #region:TAB-Generate
        private void BT_Generate_Click(object sender, EventArgs e)
        {
            #region:CreateTable 
            //instantiate datatable w/ headers
            myData = Tools.CreateDataTableHeaders(listProperties);

            //populate datatable
            foreach (Hierarchie hierarchie in listHierarchie)
            {
                Tools.PopulateDataTable(myData, hierarchie, new ArrayList());
            }
            #endregion:CreateTable

            #region:ConditionFicheEspace
            if (!RB_FE_False.Checked && RB_FE_True.Checked)
            {
                if (ficheEspaceData == null)
                {
                    MessageBox.Show("Error 404: fiche espace not found",
                        "Erreur",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }
                //Merge ficheEspace dataTable to myData
                foreach (DataRow mainDataTableRow in myData.Rows) //iterate main datatable rows
                {
                    for (int mainDataTableColumnIndex = 0; mainDataTableColumnIndex < myData.Columns.Count; mainDataTableColumnIndex++) //iterate main datatable columns w/ index
                    {
                        foreach (DataRow ficheEspaceRow in ficheEspaceData.Rows) //iterate ficheEspace datatable rows
                        {
                            if (mainDataTableRow[mainDataTableColumnIndex].ToString() == ficheEspaceRow[0].ToString())
                            {
                                foreach (DataColumn ficheEspaceColumn in ficheEspaceData.Columns)
                                {
                                    if (!myData.Columns.Contains(ficheEspaceColumn.ColumnName))
                                    {
                                        myData.Columns.Add(ficheEspaceColumn.ColumnName);
                                    }
                                    mainDataTableRow[ficheEspaceColumn.ColumnName] = ficheEspaceRow[ficheEspaceColumn.ColumnName];
                                }
                            }
                        }
                    }
                }
            }
            #endregion:ConditionFicheEspace

            #region:Export
            Excel.Worksheet destination = new Excel.Worksheet();

            if (RB_New_Excel.Checked && !RB_Existing_Excel.Checked)
            {
                Excel.Application newApp = new Excel.Application()
                {
                    Visible = true
                };
                Excel.Workbook newWb = newApp.Workbooks.Add("");
                destination = newWb.ActiveSheet;
                destination.Name = "Génération";
            }

            else if (!RB_New_Excel.Checked && RB_Existing_Excel.Checked)
            {
                destination = xlWorkbook.Worksheets.Add(After: xlWorkbook.Worksheets[xlWorkbook.Worksheets.Count]);
                destination.Activate();
                destination.Name = "Génération";
                destination.Visible = destination.Visible;
            }

            #region:Print
            bool success = false;
            if (success == false)
            {
                int colIndex = 0;
                int rowIndex = 1;

                foreach (DataColumn dc in myData.Columns)
                {
                    colIndex++;
                    destination.Cells[1, colIndex] = dc.ColumnName;
                }
                foreach (DataRow dr in myData.Rows)
                {
                    rowIndex++;
                    colIndex = 0;

                    foreach (DataColumn dc in myData.Columns)
                    {
                        colIndex++;
                        destination.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                    }
                }

                destination.Columns.AutoFit();
                success = true;
            }
            else
            {
                success = false;
            }
            #endregion:Print
            MessageBox.Show(((success) ? "Génération Réussie" : "Echec de la génération"),
                ((success) ? "Réussite" : "Erreur"),
                MessageBoxButtons.OK,
                ((success) ? MessageBoxIcon.Information : MessageBoxIcon.Error));
            #endregion:Export
        }


        #endregion:TAB-Generate
    }
}
