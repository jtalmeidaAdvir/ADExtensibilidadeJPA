
using ErpBS100;
using StdBE100;
using StdPlatBS100;
using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;

namespace ADExtensibilidadeJPA
{
    public class TrabalhadorEmpreitadaManager
    {
        private ErpBS _bso;
        private StdBSInterfPub _pso;
        private string _idEmpresa;
        private string _idEmpreitada;

        public TrabalhadorEmpreitadaManager(ErpBS bso, StdBSInterfPub pso, string idEmpresa, string idEmpreitada)
        {
            _bso = bso;
            _pso = pso;
            _idEmpresa = idEmpresa;
            _idEmpreitada = idEmpreitada;

            // Verificar se as tabelas necessárias existem
            CriarTabelasNecessarias();
        }

        private void CriarTabelasNecessarias()
        {
            try
            {
                // Verificar se a tabela de relação trabalhador-empreitada existe
                string queryCheckTable = @"
                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                               WHERE TABLE_NAME = 'TDU_AD_TrabalhadorEmpreitada')
                    BEGIN
                        CREATE TABLE [dbo].[TDU_AD_TrabalhadorEmpreitada](
                            [CDU_Id] [uniqueidentifier] NOT NULL,
                            [CDU_IdTrabalhador] [uniqueidentifier] NOT NULL,
                            [CDU_IdEmpresa] [uniqueidentifier] NOT NULL,
                            [CDU_IdEmpreitada] [nvarchar](50) NOT NULL,
                            [CDU_DataInicio] [date] NOT NULL,
                            [CDU_DataFim] [date] NULL,
                            [CDU_ContratoSubempreitada] [nvarchar](100) NULL,
                            [CDU_Status] [nvarchar](50) NOT NULL DEFAULT('Pendente'),
                            [CDU_Observacoes] [nvarchar](500) NULL,
                            CONSTRAINT [PK_TDU_AD_TrabalhadorEmpreitada] PRIMARY KEY CLUSTERED ([CDU_Id] ASC)
                        );
                    END
                ";

                _bso.DSO.ExecuteSQL(queryCheckTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao criar tabelas: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public List<Dictionary<string, object>> ObterTrabalhadoresDaEmpresa()
        {
            List<Dictionary<string, object>> trabalhadores = new List<Dictionary<string, object>>();

            try
            {
                string query = $@"SELECT 
                                    t.CDU_Id, 
                                    t.CDU_Nome, 
                                    t.CDU_TipoDocumento, 
                                    t.CDU_NumDocumento,
                                    t.CDU_NIF,
                                    t.CDU_Status,
                                    CASE WHEN te.CDU_Id IS NOT NULL THEN 1 ELSE 0 END as JaVinculado
                                FROM 
                                    TDU_AD_Trabalhadores t
                                LEFT JOIN 
                                    TDU_AD_TrabalhadorEmpreitada te ON t.CDU_Id = te.CDU_IdTrabalhador AND te.CDU_IdEmpreitada = '{_idEmpreitada}'
                                WHERE 
                                    t.CDU_IdEmpresa = '{_idEmpresa}'";

                var resultado = _bso.Consulta(query);

                if (resultado.NumLinhas() > 0)
                {
                    resultado.Inicio();
                    while (!resultado.NoFim())
                    {
                        Dictionary<string, object> trabalhador = new Dictionary<string, object>
                        {
                            { "Id", resultado.DaValor<string>("CDU_Id") },
                            { "Nome", resultado.DaValor<string>("CDU_Nome") },
                            { "TipoDocumento", resultado.DaValor<string>("CDU_TipoDocumento") },
                            { "NumDocumento", resultado.DaValor<string>("CDU_NumDocumento") },
                            { "NIF", resultado.DaValor<string>("CDU_NIF") },
                            { "Status", resultado.DaValor<string>("CDU_Status") },
                            { "JaVinculado", resultado.DaValor<int>("JaVinculado") == 1 }
                        };

                        trabalhadores.Add(trabalhador);
                        resultado.Seguinte();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao obter trabalhadores: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return trabalhadores;
        }

        public List<Dictionary<string, object>> ObterTrabalhadoresDaEmpreitada()
        {
            List<Dictionary<string, object>> trabalhadores = new List<Dictionary<string, object>>();

            try
            {
                string query = $@"SELECT 
                                    te.CDU_Id,
                                    t.CDU_Id as IdTrabalhador,
                                    t.CDU_Nome, 
                                    t.CDU_TipoDocumento, 
                                    t.CDU_NumDocumento,
                                    t.CDU_NIF,
                                    te.CDU_DataInicio,
                                    te.CDU_DataFim,
                                    te.CDU_ContratoSubempreitada,
                                    te.CDU_Status,
                                    te.CDU_Observacoes
                                FROM 
                                    TDU_AD_TrabalhadorEmpreitada te
                                INNER JOIN 
                                    TDU_AD_Trabalhadores t ON te.CDU_IdTrabalhador = t.CDU_Id
                                WHERE 
                                    te.CDU_IdEmpresa = '{_idEmpresa}'
                                    AND te.CDU_IdEmpreitada = '{_idEmpreitada}'";

                var resultado = _bso.Consulta(query);

                if (resultado.NumLinhas() > 0)
                {
                    resultado.Inicio();
                    while (!resultado.NoFim())
                    {
                        Dictionary<string, object> trabalhador = new Dictionary<string, object>
                        {
                            { "Id", resultado.DaValor<string>("CDU_Id") },
                            { "IdTrabalhador", resultado.DaValor<string>("IdTrabalhador") },
                            { "Nome", resultado.DaValor<string>("CDU_Nome") },
                            { "TipoDocumento", resultado.DaValor<string>("CDU_TipoDocumento") },
                            { "NumDocumento", resultado.DaValor<string>("CDU_NumDocumento") },
                            { "NIF", resultado.DaValor<string>("CDU_NIF") },
                            { "DataInicio", resultado.DaValor<DateTime>("CDU_DataInicio") },
                            { "DataFim", resultado.DaValor<DateTime?>("CDU_DataFim") },
                            { "ContratoSubempreitada", resultado.DaValor<string>("CDU_ContratoSubempreitada") },
                            { "Status", resultado.DaValor<string>("CDU_Status") },
                            { "Observacoes", resultado.DaValor<string>("CDU_Observacoes") }
                        };

                        trabalhadores.Add(trabalhador);
                        resultado.Seguinte();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao obter trabalhadores da empreitada: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return trabalhadores;
        }

        public bool VincularTrabalhadorEmpreitada(string idTrabalhador, DateTime dataInicio, DateTime? dataFim, string contratoSubempreitada, string status, string observacoes)
        {
            try
            {
                // Verificar se o trabalhador já está vinculado a esta empreitada
                string queryVerificar = $@"
                    SELECT COUNT(*) AS Total 
                    FROM TDU_AD_TrabalhadorEmpreitada 
                    WHERE CDU_IdTrabalhador = '{idTrabalhador}' 
                    AND CDU_IdEmpreitada = '{_idEmpreitada}'";

                var resultado = _bso.Consulta(queryVerificar);
                resultado.Inicio();
                int total = resultado.DaValor<int>("Total");

                string id = Guid.NewGuid().ToString();
                string dataFimFormatada = dataFim.HasValue ? $"'{dataFim.Value:yyyy-MM-dd}'" : "NULL";

                if (total > 0)
                {
                    // Atualizar vínculo existente
                    string queryAtualizar = $@"
                        UPDATE TDU_AD_TrabalhadorEmpreitada SET
                            CDU_DataInicio = '{dataInicio:yyyy-MM-dd}',
                            CDU_DataFim = {dataFimFormatada},
                            CDU_ContratoSubempreitada = '{contratoSubempreitada?.Replace("'", "''")}',
                            CDU_Status = '{status}',
                            CDU_Observacoes = '{observacoes?.Replace("'", "''")}'
                        WHERE CDU_IdTrabalhador = '{idTrabalhador}' 
                        AND CDU_IdEmpreitada = '{_idEmpreitada}'";

                    _bso.DSO.ExecuteSQL(queryAtualizar);
                }
                else
                {
                    // Criar novo vínculo
                    string queryInserir = $@"
                        INSERT INTO TDU_AD_TrabalhadorEmpreitada (
                            CDU_Id, CDU_IdTrabalhador, CDU_IdEmpresa, CDU_IdEmpreitada,
                            CDU_DataInicio, CDU_DataFim, CDU_ContratoSubempreitada, CDU_Status, CDU_Observacoes
                        ) VALUES (
                            '{id}', '{idTrabalhador}', '{_idEmpresa}', '{_idEmpreitada}',
                            '{dataInicio:yyyy-MM-dd}', {dataFimFormatada}, '{contratoSubempreitada?.Replace("'", "''")}',
                            '{status}', '{observacoes?.Replace("'", "''")}'
                        )";

                    _bso.DSO.ExecuteSQL(queryInserir);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao vincular trabalhador à empreitada: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public bool RemoverTrabalhadorEmpreitada(string idVinculo)
        {
            try
            {
                string queryRemover = $@"DELETE FROM TDU_AD_TrabalhadorEmpreitada WHERE CDU_Id = '{idVinculo}'";
                _bso.DSO.ExecuteSQL(queryRemover);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao remover trabalhador da empreitada: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public bool AtualizarStatusTrabalhadorEmpreitada(string idVinculo, string novoStatus, string observacoes)
        {
            try
            {
                string queryAtualizar = $@"
                    UPDATE TDU_AD_TrabalhadorEmpreitada SET
                        CDU_Status = '{novoStatus}',
                        CDU_Observacoes = '{observacoes?.Replace("'", "''")}'
                    WHERE CDU_Id = '{idVinculo}'";

                _bso.DSO.ExecuteSQL(queryAtualizar);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar status do trabalhador na empreitada: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
