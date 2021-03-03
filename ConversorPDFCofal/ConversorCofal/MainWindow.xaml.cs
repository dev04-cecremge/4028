using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using IO = System.IO;
using Microsoft.Win32;
using ConversorCofal.Helpers;
using System.Diagnostics;

namespace ConversorCofal
{
    /// <summary>
    /// Interação lógica para MainWindow.xam
    /// </summary>
    public partial class MainWindow : Window
    {

        PDF pdf = new PDF();
        EXCEL excel = new EXCEL();

        public void Fechar(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        //Apaga todos os forms que aparecem durante a exibição do relatório
        public void Apaga(bool inst, bool passo1, bool passo2, bool passo3)
        {
            if (inst) instrucoes.Visibility = Visibility.Visible; else instrucoes.Visibility = Visibility.Collapsed;

            if (passo1) {
                passo1Button.Visibility = Visibility.Visible; passo1Txt.Visibility = Visibility.Visible;
            }
            else{
                passo1Button.Visibility = Visibility.Collapsed; passo1Txt.Visibility = Visibility.Collapsed;
            }

            if (passo2)
            {
                passo2Button.Visibility = Visibility.Visible; passo2Txt.Visibility = Visibility.Visible;
            }
            else
            {
                passo2Button.Visibility = Visibility.Collapsed; passo2Txt.Visibility = Visibility.Collapsed;
            }

            if(passo3){
                passo3Txt.Visibility = Visibility.Visible; passo3Button.Visibility = Visibility.Visible;
            }
            else
            {
                passo3Txt.Visibility = Visibility.Collapsed; passo3Button.Visibility = Visibility.Collapsed;
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            //Zera as instruções e os 3passos
            Apaga(false, false, false, false);

        }

        private void MenuItemRelatorioAposentados_Click(object sender, RoutedEventArgs e)
        {
            //Exibir instruções e passo 1
            Apaga(true, true, false, false);
        }

        //Clicou no passo 1 - Excel
        private void SelecionarExcel_Click(object sender, RoutedEventArgs e)
        {
            //Apago mensagens de erro
            textLblErrorEXCEL.Text = "";

            //Cria instancia
            OpenFileDialog arquivo = new OpenFileDialog();

            //Pode vir True, False ou NULL
            bool? ret;
            ret= arquivo.ShowDialog();

            //Arquivo foi selecionado
            if (ret == true)
            {

                //Validar Excel
                //1 - Abrir XLSX
                string path = arquivo.InitialDirectory + arquivo.FileName;

                int diag = excel.CarregarExcel(path);
                //1 -  Sucesso
                //-1 - Sem path
                //-2 - Nao pode abrir o documento
                textLblErrorEXCEL.Text = diag == -1 ? textLblErrorEXCEL.Text = "Favor especificar um local válido" : textLblErrorEXCEL.Text = "";
                textLblErrorEXCEL.Text = diag == -2 ? textLblErrorEXCEL.Text = "Não consegui abrir o Excel. Erro devido ao local, formato ou permissão!" : textLblErrorEXCEL.Text = "";
                //Verifica se existe erro nas regras do excel
                if (diag < 0) return;
                
                //Validando regras de Excel
                int excelValido = excel.ValidarRegrasExcel();
                if (excelValido == -5) { textLblErrorEXCEL.Text = "Erro no cabecalho. O mesmo deve conter Nome Cliente e na segunda célula CPF/CNPJ"; return; }
                if (excelValido == -6) { textLblErrorEXCEL.Text = "Erro em um ou mais nomes. foi identifica um nome com caracteres especiais ou números"; return; }
                if (excelValido == -7) { textLblErrorEXCEL.Text = "Erro em um ou mais CPFS. foi identifica um CPF com caracteres que não são numéricos"; return; }

                
                //Habilito o passo 2
                Apaga(true, true, true, false);
                
            }
            //Deu errado
            else
            {
                //Desabilito os passos 2 e 3
                Apaga(true, true, false, false);

            }


        }

        //Clicou no passo 2 - PDF
        private void SelecionarPdf_Click(object sender, RoutedEventArgs e)
        {
            //Remover texto de erro
            textLblErrorPDF.Text = "";

            //Cria instancia
            OpenFileDialog arquivo = new OpenFileDialog();

            //Pode vir True, False ou NULL
            bool? ret;
            ret = arquivo.ShowDialog();

            //Arquivo foi selecionado
            if (ret == true)
            {

                //--------Validar PDF
                //1 - Abrir PDF
                string path = arquivo.InitialDirectory + arquivo.FileName;

                int diag = pdf.CarregarPdf(path);
                //1 -  Sucesso
                //-1 - Sem path
                //-2 - Nao pode abrir o documento
                textLblErrorPDF.Text = diag == -1 ? textLblErrorPDF.Text = "Favor especificar um local valido" : textLblErrorPDF.Text = "";
                textLblErrorPDF.Text = diag == -2 ? textLblErrorPDF.Text = "Não consegui abrir o arquivo. Erro devido ao local, formato ou permissão!" : textLblErrorPDF.Text = "";

                //Habilito os novos campos ou nao?
                if (diag==1)
                {
                    //Habilito o passo 3
                    Apaga(true, true, true, true);
                }

            }
            //Deu errado
            else
            {
                Apaga(true, true, true, false);
            }


        }

        //Clicou no passo 3 - Gerar Excel
        private void GerarPDF_CLick(object sender, RoutedEventArgs e)
        {

            //Edita conforme especificado no Excel
            //Cortar PDF
            pdf.CortarPaginasPDF(excel);

            //Mostra erro que significa que o PDF gerado nao tem paginas
            if (pdf.arquivo.PageCount > 0)
            {
                //Salva no %temp%
                string output = IO.Path.GetTempPath() + "resultado.pdf";
                pdf.SalvarPdf(output);

                //Abre o arquivo¨s
                try
                {
                    pdf.AbrirPdf();
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                //Apaga todos os campos 
                Apaga(false, false, false, false);

            }
            else
            {
                //Erro de arquivo gerado sem resultados.
                textLblErrorGerarPDF.Text = "Opa. Parece que nenhum CPF da planilha existe no PDF final... verifique os mesmos e tente novamente!";
            }
            
        }

    }
}
