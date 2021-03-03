VERSÃO 1.0
VISUAL STUDIO: 2017

1 - O que faz:
	1.1 - Importa um excel contendo nome/cpf de todos os associados do relatório de imposto de renda
	1.2 - Importa o relatório de imposto de renda completo gerado pelo SISBR
		1.3 - Gera um novo PDF contendo apenas os associados da planilha

2 - Necessidade:
	2.1 - O SISBR não permite executar filtros no relatório de imposto de renda. 
		Logo, esse sistema realiza o "corte" no PDF gerado pelo SISBR, tendo como 
		saida apenas o as páginas que cujos associados estão na planilha.

3 - Dependências:
	3.1 - Feito em WPF (https://docs.microsoft.com/en-us/dotnet/desktop/wpf/?view=netdesktop-5.0)
	3.2 - EStrutura: .Net Framework 4.6.1
	3.3 - Usa excel, a partir da EEPLUS 4.5.3.2
	3.4 - Usa PDF, a partir do PDFSharp 1.50.5147.0

4 - Estrutura:
	4.1 - MainWindow.xaml - Posuui todas as implementações de tela e do relatório.
	4.2 - MainWindow.xaml/MainWindow.xaml.cs - Contém todo o código por trás em C#.
	4.3 - Assets - Contém imagens e ícones usados no sistema
	4.4 - Helppers - Contém as classes para utilização do PDF e EXCEL