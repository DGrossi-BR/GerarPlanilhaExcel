using System;
using System.Data.SqlClient;
using System.IO;

namespace GerarPlanilhaExcel
{
    class Program
    {
        //  Cria uma variável com o pacote SqlConnection
        public static SqlConnection conexao;

        static void Main()
        {
            Console.WriteLine("Conectando com o banco de dados...");

            //  Cria a conexão com a base de dados
            conexao = new SqlConnection("Server=.;Database=Agencia_Viagem_ADM;Trusted_Connection=True;MultipleActiveResultSets=true;");
            conexao.Open();

            Console.WriteLine("Definindo caminho que será salvo o arquivo...");

            //  Define o caminho onde o arquivo será gravado
            string caminho = "C:/Teste/ArquivoExcelQueFoiGerado.xls";

            Console.WriteLine("Criando arquivo...");

            try
            {
                //  Cria o arquivo Excel que vai conter os dados da tabela passando o caminho configurado anteriormente
                using (StreamWriter arquivoExcel = File.CreateText(caminho))
                {
                    //  Insere os nomes das colunas no arquivo Excel criado
                    arquivoExcel.WriteLine("IdEmpresa" + "\t" + "RazaoSocial" + "\t" + "NomeFantasia" + "\t" + "CNPJ");

                    Console.WriteLine("Buscando os dados para inserir no arquivo criado...");

                    //  Faz a consulta no banco de dados
                    var selectDadosTabela = conexao.CreateCommand();
                    selectDadosTabela.CommandText = "SELECT * FROM Empresa";
                    var resultadoDadosDaTabela = selectDadosTabela.ExecuteReader();

                    Console.WriteLine("Inserindo os dados no arquivo...");

                    //  Insere o resultado da consulta no arquivo Excel criado
                    while (resultadoDadosDaTabela.Read())
                    {
                        arquivoExcel.WriteLine(resultadoDadosDaTabela["IdEmpresa"].ToString() + "\t" + resultadoDadosDaTabela["RazaoSocial"].ToString() + "\t" + resultadoDadosDaTabela["NomeFantasia"].ToString() + "\t" + resultadoDadosDaTabela["CNPJ"].ToString());
                    }

                    Console.WriteLine("Fechando conexão com o banco de dados...");

                    //  Fecha a conexão com o banco de dados
                    conexao.Dispose();

                    Console.WriteLine("Arquivo gerado com sucesso no caminho: " + caminho);
                    Console.WriteLine("Pressione qualquer tecla para finalizar o processo.");
                    Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro ao tentar gerar o arquivo. Segue maiores detalhes do erro: " + ex.Message);
                Console.WriteLine("Pressione qualquer tecla para finalizar o processo.");
                Console.ReadKey();
            }
        }
    }
}
