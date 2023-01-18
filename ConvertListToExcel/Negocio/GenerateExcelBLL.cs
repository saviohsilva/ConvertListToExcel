using ConvertListToExcel.Configuracoes;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Drawing;
using System.Reflection;

//https://www.blogson.com.br/gerar-uma-planilha-do-excel-com-c/

namespace ConvertListToExcel.Negocio
{
  public static class GenerateExcelBLL
  {
    /// <summary>
    /// Transforma uma lista de objetos em um arquivo do excel.
    /// </summary>
    /// <param name="lstItens">Lista de objetos a serem inseridos na planilha.</param>
    /// <param name="configCabecalho">Configuração de formatação das células do cabeçalho</param>
    /// <param name="configCorpo">Configuração de formatação das células do corpo da planilha</param>
    /// <param name="salvarPlanilhaConfig">Classe de configuração com as configurações do salvamento da planilha</param>
    public static void GerarPlanilha(IList lstItens, EstiloFonteConfig configCabecalho, EstiloFonteConfig configCorpo, SalvarPlanilhaConfig salvarPlanilhaConfig)
    {
      try
      {
        //Inicializa as propriedades do excel e cria uma planilha temporária em memória.
        object misValue = Missing.Value;
        Application xlApp = new Application();
        Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
        Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

        int contLinha = 1; //Contador de linhas.
        int contColuna = 1; //Contador de colunas. 

        //Verifica se o diratório é válido.
        if (!Directory.Exists(salvarPlanilhaConfig.Diretorio))
          throw new Exception(ConvertListToExcelResource.DiretorioInvalido);

        //Para transformar a lista em planilha é necessário que a mesma possua apenas um tipo de objetos.
        if (lstItens.GetType().GenericTypeArguments.Count() > 1)
          throw new Exception(ConvertListToExcelResource.FalhaListaComMaisTipos);

        //Recupera o tipo de objetos contidos na lista
        Type type = lstItens.GetType().GenericTypeArguments.First();
        PropertyInfo[] props = type.GetProperties();

        //Cria o cabeçalho com o nome das propriedades. 
        foreach (var prop in props)
        {
          PreencheCelula(ref contLinha, ref contColuna, configCabecalho, xlWorkSheet, prop.Name);
          contColuna++;
        }

        //Começa a percorrer os itens da lista preenchendo-os nas linhas da planilhas.
        for (int i = 0; i < lstItens.Count; i++)
        {
          //Sempre que um novo item for adicionado na planilha, é necessário incrementar no contador de linhas e zerar o contador de colunas.
          contLinha++;
          contColuna = 1;

          var item = lstItens[i];

          foreach (var prop in props)
          {
            //Caso a propriedade não for primitiva, ela não será criada na tabela. 
            if (prop.PropertyType.IsArray || prop is IList || prop.PropertyType.IsCollectible)
              continue;

            PreencheCelula(ref contLinha, ref contColuna, configCorpo, xlWorkSheet, prop.GetValue(item));

            contColuna++;
          }

          //Atribui tamanho automático para todas as colunas da planilha. 
          xlWorkSheet.Columns.AutoFit();
        }

        SalvarPlanilha(xlWorkBook, xlApp, misValue, salvarPlanilhaConfig);
      }
      catch (Exception ex)
      {
        Console.WriteLine(ConvertListToExcelResource.ErroGerarPlanilha, ex);
      }
    }

    /// <summary>
    /// Método responsável pelo preenchimento das células da planilha. 
    /// </summary>
    private static void PreencheCelula(ref int ordemLinha, ref int ordemColuna, EstiloFonteConfig estiloCelula, Worksheet xlWorkSheet, dynamic? valor)
    {
      xlWorkSheet.Cells[ordemLinha, ordemColuna] = valor;
      FormataCelula(xlWorkSheet, ordemLinha, ordemColuna, estiloCelula);
    }

    /// <summary>
    /// Método responsável por salvar a planilha.
    /// </summary>
    private static void SalvarPlanilha(Workbook xlWorkBook, Application xlApp, object misValue, SalvarPlanilhaConfig salvarPlanilhaConfig)
    {
      //Salva o arquivo de acordo com a documentação do Excel.
      xlWorkBook.SaveAs(string.Format("{0}\\{1}.xlsx", salvarPlanilhaConfig.Diretorio, salvarPlanilhaConfig.Nome),
                        XlFileFormat.xlOpenXMLStrictWorkbook,
                        string.IsNullOrEmpty(salvarPlanilhaConfig.Senha) ? misValue : salvarPlanilhaConfig.Senha,
                        misValue,
                        misValue,
                        misValue,
                        XlSaveAsAccessMode.xlExclusive,
                        misValue,
                        misValue,
                        misValue,
                        misValue,
                        misValue);

      xlWorkBook.Close(true, misValue, misValue);
      xlApp.Quit();
    }

    /// <summary>
    /// Método responsável por realizar a formatação das células. 
    /// </summary>
    private static void FormataCelula(Worksheet xlWorkSheet, int ordemLinha, int ordemColuna, EstiloFonteConfig estiloCelula)
    {
      //Aplica customizações nas celulas. 
      if (estiloCelula.Bordas)
        xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Cells.Borders.LineStyle = XlLineStyle.xlContinuous; //Bordas.

      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Font.Size = estiloCelula.TamanhoFonte; //Tamanho da fonte.
      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Font.FontStyle = estiloCelula.FonteStyle; //Stilo da fonte.
      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Font.Color = ColorTranslator.ToOle(Color.FromName(estiloCelula.CorFonte));  //Cro da fonte.
      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Font.Bold = estiloCelula.FonteNegrito; //Fonte em negrito.
      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Font.Italic = estiloCelula.FonteItalica; //Fonte em itálico.
      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Font.Underline = estiloCelula.FonteSublinhada; //Fonte com sublinhado.
      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).HorizontalAlignment = estiloCelula.AlinhamentoHorizontal; //Define o alinhamento horizontal.
      xlWorkSheet.get_Range(RecuperarNomeColuna(ordemLinha, ordemColuna), RecuperarNomeColuna(ordemLinha, ordemColuna)).Interior.Color = ColorTranslator.ToOle(Color.FromName(estiloCelula.CorFundo));  //Cro de fundo da celula.
    }

    /// <summary>
    /// Método responsável por realizar a conversão do número de coluna para o nome da mesma.
    /// </summary>
    private static string RecuperarNomeColuna(int numeroLinha, int numeroColuna)
    {
      int divisor = numeroColuna;
      string nomeColuna = string.Empty;
      int modulo;

      while (divisor > 0)
      {
        modulo = (divisor - 1) % 26;
        nomeColuna = Convert.ToChar(65 + modulo).ToString() + nomeColuna;
        divisor = (divisor - modulo) / 26;
      }

      return string.Format("{0}{1}", nomeColuna, numeroLinha);
    }
  }
}
