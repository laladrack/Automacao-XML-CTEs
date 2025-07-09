using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automação_XML_CTEs
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void dateTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            dateTextBox.Text = DateTime.Now.ToString("dd/MM/yyyy hh:mm");
        }

        private void folderSelectButton_Click(object sender, RoutedEventArgs e)
        {
            string XMLsFolder;
            var folderDialog = new OpenFolderDialog
            {
                Title = "Selecione a pasta com os XMLs.",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (folderDialog.ShowDialog() == true)
            {
                XMLsFolder = folderDialog.FolderName;
                folderSelectLabel.Content = $"{XMLsFolder}";
            }
        }

        private void savingFolderButton_Click(object sender, RoutedEventArgs e)
        {
            string saveFolder;
            var folderDialog = new OpenFolderDialog
            {
                Title = "Selecione a pasta para salvar o relatório.",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (folderDialog.ShowDialog() == true)
            {
                saveFolder = folderDialog.FolderName;
                savingFolderLabel.Content = $"{saveFolder}";
            }
        }

        private void gerarRelatorioButton_Click(object sender, RoutedEventArgs e)
        {
            //declarar caminhos
            string XMLsFolder = folderSelectLabel.Content.ToString();
            string savingFolder = savingFolderLabel.Content.ToString();
            string nomeReport = labelNomeReport.Text + ".xlsx";
            string filePath = "-";
            string emNome = "-";
            string emCNPJ = "-";
            string vComp = "-";
            string vCarga = "-";
            string UFEnv = "-";
            string pICMS = "-";
            string tomaNome = "-";
            string tomaCNPJ = "-";
            string nCT = "-";
            string dhEmi = "-";

            //criar o report
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook relatorio = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Worksheet)relatorio.ActiveSheet;

            //cabeçalho
            worksheet.Range["A1"].Value = "Filename";
            worksheet.Range["B1"].Value = "NCT";
            worksheet.Range["C1"].Value = "Data de emissão";
            worksheet.Range["D1"].Value = "Valor de carga";
            worksheet.Range["E1"].Value = "Alíquota ICMS";
            worksheet.Range["F1"].Value = "Pedágio";
            worksheet.Range["G1"].Value = "Estado do remetente";
            worksheet.Range["H1"].Value = "CNPJ emissor";
            worksheet.Range["I1"].Value = "Nome emissor";
            worksheet.Range["J1"].Value = "CNPJ tomador";
            worksheet.Range["K1"].Value = "Nome tomador";

            int i = 1;
            //loop xml por xml
            if (XMLsFolder != "Nenhuma pasta selecionada" && savingFolder != "Nenhuma pasta selecionada")
            {
                string[] files = Directory.GetFiles(XMLsFolder);
                foreach (string file in files)
                {
                    if (file.EndsWith(".xml") == true)
                    {
                        //filePath = System.IO.Path.Combine(XMLsFolder, file);

                        XDocument xml = XDocument.Load(file);
                        XElement infCte = xml.Element("cteProc").Element("{http://www.portalfiscal.inf.br/cte}CTe").Element("{http://www.portalfiscal.inf.br/cte}infCte");

                        //nct
                        nCT = infCte.Element("{http://www.portalfiscal.inf.br/cte}ide").Element("{http://www.portalfiscal.inf.br/cte}nCT").Value.ToString();
                        //dhEmi
                        dhEmi = infCte.Element("{http://www.portalfiscal.inf.br/cte}ide").Element("{http://www.portalfiscal.inf.br/cte}dhEmi").Value.ToString();
                        //tomador
                        if (infCte.Element("{http://www.portalfiscal.inf.br/cte}ide").Element("{http://www.portalfiscal.inf.br/cte}toma4") == null)
                        {
                            switch (infCte.Element("{http://www.portalfiscal.inf.br/cte}ide").Element("{http://www.portalfiscal.inf.br/cte}toma3").Element("{http://www.portalfiscal.inf.br/cte}toma").Value.ToString())
                            {
                                case "0": //remetente
                                    tomaCNPJ = infCte.Element("{http://www.portalfiscal.inf.br/cte}rem").Element("{http://www.portalfiscal.inf.br/cte}CNPJ").Value.ToString();
                                    tomaNome = infCte.Element("{http://www.portalfiscal.inf.br/cte}rem").Element("{http://www.portalfiscal.inf.br/cte}xNome").Value.ToString();
                                    break;

                                case "1": //expedidor
                                    tomaCNPJ = infCte.Element("{http://www.portalfiscal.inf.br/cte}exped").Element("{http://www.portalfiscal.inf.br/cte}CNPJ").Value.ToString();
                                    tomaNome = infCte.Element("{http://www.portalfiscal.inf.br/cte}exped").Element("{http://www.portalfiscal.inf.br/cte}xNome").Value.ToString();
                                    break;

                                case "2": //recebedor
                                    tomaCNPJ = infCte.Element("{http://www.portalfiscal.inf.br/cte}receb").Element("{http://www.portalfiscal.inf.br/cte}CNPJ").Value.ToString();
                                    tomaNome = infCte.Element("{http://www.portalfiscal.inf.br/cte}receb").Element("{http://www.portalfiscal.inf.br/cte}xNome").Value.ToString();
                                    break;

                                case "3": //destinatário
                                    tomaCNPJ = infCte.Element("{http://www.portalfiscal.inf.br/cte}dest").Element("{http://www.portalfiscal.inf.br/cte}CNPJ").Value.ToString();
                                    tomaNome = infCte.Element("{http://www.portalfiscal.inf.br/cte}dest").Element("{http://www.portalfiscal.inf.br/cte}xNome").Value.ToString();
                                    break;
                            }
                        }
                        else
                        {
                            tomaCNPJ = infCte.Element("{http://www.portalfiscal.inf.br/cte}ide").Element("{http://www.portalfiscal.inf.br/cte}toma4").Element("{http://www.portalfiscal.inf.br/cte}CNPJ").Value.ToString();
                            tomaNome = infCte.Element("{http://www.portalfiscal.inf.br/cte}ide").Element("{http://www.portalfiscal.inf.br/cte}toma4").Element("{http://www.portalfiscal.inf.br/cte}xNome").Value.ToString();
                        }

                        //emit
                        emNome = infCte.Element("{http://www.portalfiscal.inf.br/cte}emit").Element("{http://www.portalfiscal.inf.br/cte}xNome").Value.ToString();
                        emCNPJ = infCte.Element("{http://www.portalfiscal.inf.br/cte}emit").Element("{http://www.portalfiscal.inf.br/cte}CNPJ").Value.ToString();

                        //pedagio

                        string textoTag = infCte.Element("{http://www.portalfiscal.inf.br/cte}vPrest").Element("{http://www.portalfiscal.inf.br/cte}Comp").Element("{http://www.portalfiscal.inf.br/cte}xNome").Value.ToString();
                        if (textoTag.ToLower() == "pedagio" || textoTag.ToLower() == "pedágio")
                        {
                            vComp = infCte.Element("{http://www.portalfiscal.inf.br/cte}vPrest").Element("{http://www.portalfiscal.inf.br/cte}Comp").Element("{http://www.portalfiscal.inf.br/cte}vComp").Value.ToString();
                        }
                        foreach (var elemento in infCte.Elements())
                        {
                            if (elemento.Name.ToString().Contains("vCarga"))
                            {
                                vCarga = elemento.Value.ToString();
                            }
                            else if (elemento.Name.ToString().Contains("UFEnv"))
                            {
                                UFEnv = elemento.Value.ToString();
                            }
                            else if (elemento.Name.ToString().Contains("pICMS"))
                            {
                                pICMS = elemento.Value.ToString();
                            }
                        }

                        //criar report
                        worksheet.Cells[i, 1].Value = file;
                        worksheet.Cells[i, 2].Value = nCT;
                        worksheet.Cells[i, 3].Value = dhEmi;
                        worksheet.Cells[i, 4].Value = vCarga;
                        worksheet.Cells[i, 5].Value = pICMS;
                        worksheet.Cells[i, 6].Value = vComp;
                        worksheet.Cells[i, 7].Value = UFEnv;
                        worksheet.Cells[i, 8].Value = emCNPJ;
                        worksheet.Cells[i, 9].Value = emNome;
                        worksheet.Cells[i, 10].Value = tomaCNPJ;
                        worksheet.Cells[i, 11].Value = tomaNome;

                        i++;

                        relatorio.SaveAs2(System.IO.Path.Combine(savingFolder, nomeReport));
                        gerarRelatorioLabel.Content = "Relatório salvo em " + savingFolder;
                    }
                }
            }
        }
    }
}
