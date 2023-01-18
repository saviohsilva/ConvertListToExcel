using ConvertListToExcel.Configuracoes;
using ConvertListToExcel.Modelos;
using ConvertListToExcel.Negocio;

var pessoas = new List<PessoaModel>()
{
  new PessoaModel()
  {
    Nome = "Fulana",
    Endereco = "Rua 02",
    Idade = 10,
    sexo = "Feminino"
  },

  new PessoaModel()
  {
    Nome = "Alguem",
    Endereco = "Rua S/N",
    Idade = 100,
    sexo = "Masculino"
  }
};

var configCabecalho = new EstiloFonteConfig()
{
  TamanhoFonte = 16,
  CorFonte = "Red",
  CorFundo = "Gray",
  FonteSublinhada = true,
  FonteNegrito = true,
};

var configCorpo = new EstiloFonteConfig();

var configSalvar = new SalvarPlanilhaConfig()
{
  Diretorio = @"C:\Temp",
  Nome = "Planilha Teste",
  Senha ="123",
};

GenerateExcelBLL.GerarPlanilha(pessoas, configCabecalho, configCorpo, configSalvar);