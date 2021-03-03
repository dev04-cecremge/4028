using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConversorCofal.Helpers
{
    public class PDF
    {
        string pathComFormato = String.Empty;
        string pathSaidaComFormato = String.Empty;
        public PdfDocument arquivo = new PdfDocument();

        //1 -  Sucesso
        //-1 - Sem path
        //-2 - Nao pode abrir o documento
        public int CarregarPdf(string _pathComFormato)
        {
            pathComFormato = _pathComFormato;

            if (pathComFormato.Length == 0)
            {
                return -1;
            }

            try
            {
                arquivo = PdfReader.Open(pathComFormato);
            }
            catch (Exception ex)
            {
                return -2;
            }

            //Sucesso
            return 1;

        }

        //1 - Sucesso
        //-1 - Arquivo nao importado
        //-2 - Erro ao salvar no local
        //-3 - Path nao informado
        public int SalvarPdf(string _pathComFormato)
        {
            //Ver se existe um documento carregado
            if (!arquivo.IsImported) return -1;
            //Sem local de saida:
            if (_pathComFormato.Length < 1) return -3;

            //Salvo path de saida
            pathSaidaComFormato = _pathComFormato;

            try
            {
                arquivo.Save(_pathComFormato);

            }catch(Exception ex){
                return -2;
            }

            //Sucess
            return 1;
        }

        //True - Deu certo
        //False - Deu errado
        public bool AbrirPdf()
        {
            try
            {
                Process.Start(@"" + pathSaidaComFormato);
                return true;
            }
            catch(Exception ex){
                return false;
            }
            
        }

        //Relatorio que varia conforme excel
        public bool CortarPaginasPDF(EXCEL excel)
        {
            //Abrir os CPFS em lista
            List<string> cpfs = excel.listaCPFs;

            //cortar primiera pagina
            int numeroPaginas = arquivo.PageCount;

            for (int pagina = 0; pagina < numeroPaginas; pagina++)
            {

                //Pego todo os texto ou pode ter chegado no fim, porque o numero maximo diminui!

                string todoTexto = arquivo.Pages[pagina].Contents.Elements.GetDictionary(0).Stream.ToString();

                string cpfAtual = todoTexto.Substring( Regex.Match(todoTexto, @"\d{3}.\d{3}.\d{3}\-\d{2}").Index , 14);
                //formata cpfAtual
                cpfAtual = cpfAtual.Remove(11, 1);
                cpfAtual = cpfAtual.Remove(7, 1);
                cpfAtual = cpfAtual.Remove(3, 1);


                //Verificar se o CPF da pessoa dessa pagina existe na relacao do excel
                if (cpfs.Contains(cpfAtual)) {
                    //Nada
                }
                else
                {
                    arquivo.Pages.RemoveAt(pagina); pagina--; numeroPaginas--;
                }

            }


            return true;
        }

        public bool RemoverPaginaX(int x)
        {
            if (x < 0 || x > arquivo.PageCount) return false;

            try
            {
                arquivo.Pages.RemoveAt(x);
                return true;

            }catch(Exception ex)
            {
                throw ex;
            }
            
        }


    }
}
