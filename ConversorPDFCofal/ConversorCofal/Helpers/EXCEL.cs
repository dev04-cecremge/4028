using OfficeOpenXml; //EPPLUS
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace ConversorCofal.Helpers
{
    public class EXCEL
    {
        //Apenas o pacote, dentro dele existe um elemento qu guarda a planilha e dentro dele, a planilha
        ExcelPackage pacote = new ExcelPackage();

        string pathEntradaComExtensao = String.Empty;
        string pathSaidaComExtensao = String.Empty;

        public List<string> listaCPFs = new List<string>();

        //1 -  Sucesso
        //-1 - Sem path
        //-2 - Nao pode abrir o documento
        public int CarregarExcel(string _pathComExtensao)
        {
            pathEntradaComExtensao = _pathComExtensao;

            //Erro -1
            if (pathEntradaComExtensao.Length < 1) return -1;

            try
            {
                pacote = new ExcelPackage( new System.IO.FileInfo(pathEntradaComExtensao) );
                return 1;
            }
            catch (Exception ex)
            {
                return -2;
            }

        }

        //-5 - A primeira celula deve conter a frase "Nome Cliente" e na segunda celula: "CPF/CNPJ"
        //-6 - Um dos nomes tem números o ucaracterres especiais
        //-7 - um dos CPFS ou CPNJs contem formato invalido. DEVEM apenas ter numeros
        public int ValidarRegrasExcel()
        {
            ExcelWorkbook workbook = pacote.Workbook;

            //Primeira ABA
            ExcelWorksheet abaDaPlanilha = workbook.Worksheets[1];

            //Erro de primeira linha
            bool ok;
            ok= abaDaPlanilha.Cells[1, 1].Text == "Nome Cliente" ? true : false;
            ok = abaDaPlanilha.Cells[1, 2].Text == "CPF/CNPJ" ? true : false;
            if (!ok) return -5;

            //Erro de nomes e de CPFS
            var start = abaDaPlanilha.Dimension.Start;
            var end = abaDaPlanilha.Dimension.End;

            //Row+1 para tirar o cabecalho
            for (int row = start.Row+1; row <= end.Row; row++)
            {
                //Nome
                //Ver se nome contem somente letras e espacos.
                if (!Regex.IsMatch(abaDaPlanilha.Cells[row, 1].Text, @"^[a-zA-Z\s]+$"))
                {
                    var r = row;
                    var val = abaDaPlanilha.Cells[row, 1].Text;
                    return -6;
                }

                //CPF
                //Ver se um dos CPFS contem algo alem de numeros
                if (!Regex.IsMatch(abaDaPlanilha.Cells[row, 2].Text, @"^[\d]+$")) return -7;

                //Inserir esse CPF na lista de CPFS
                listaCPFs.Add( abaDaPlanilha.Cells[row, 2].Text);

            }

            //Sucesso
            return 1;
        }

        public bool SalvarExcel()
        {
            return true;
        }

    }
}
