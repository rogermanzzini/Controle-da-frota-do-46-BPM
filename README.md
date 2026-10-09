# 46º BPM (Bagé / RS) - Sistema de Relatório de Consumo e Abastecimento de Viaturas

Sistema desenvolvido para o **46º Batalhão de Polícia Militar (Bagé/RS)** para processar planilhas de abastecimentos estaduais do Rio Grande do Sul (RS), cruzar estritamente com os veículos de interesse da unidade e calcular com precisão o consumo ($km/l$), quilometragem percorrida, custos e inconsistências de odômetro.

---

## 🚀 Tecnologias Utilizadas

- **Processamento de Dados**: `pandas` + `openpyxl` (limpeza de caracteres, normalização de placas/prefixos, cálculo temporal de intervalos e geração de planilhas formatadas).
- **Banco de Dados**: `SQLite` (`bpm46_combustivel.db` - persistência de frota, lotes de abastecimentos e indicadores consolidados).
- **Backend**: `Python 3.12` + `Flask` (rotas RESTful, upload de arquivos, regras de negócio e controle de sessões).
- **Frontend**: HTML5, CSS3 Tático Militar, JavaScript moderno, Bootstrap 5 e `Chart.js` para visualização gráfica.

---

## 📋 Como Funciona o Cruzamento de Dados

1. **Planilha Geral do RS**:
   - Pode conter milhares de linhas de abastecimentos de viaturas de diversos batalhões, cidades e órgãos de segurança de todo o estado do RS.
2. **Planilha de Viaturas Alvo**:
   - O operador fornece uma planilha contendo apenas as viaturas do **46º BPM (Bagé)** com Placas e/ou Prefixos (ex: Duster, Hilux Força Tática, Spin, Corolla).
   - O sistema também permite salvar a frota no SQLite e utilizá-la em importações futuras sem precisar reenviar a lista de viaturas.
3. **Cálculos Realizados automaticamente via Pandas**:
   - **$\Delta km$**: Diferença entre odômetro do abastecimento atual e o imediatamente anterior da mesma viatura.
   - **Consumo ($km/l$)**: $\frac{\Delta km}{\text{litros abastecidos}}$.
   - **Consumo Médio Consolidado**: Total de quilômetros úteis dividido pelo total de litros consumidos no período.
   - **Custo por Quilômetro**: $\frac{\text{Valor Gasto (R\$)}}{\Delta km}$.
   - **Detecção de Anomalias**: Alertas de odômetro regressivo, registros duplicados e médias atípicas ($< 2.5$ ou $> 30.0\ km/l$).

---

## 🏃 Como Executar

### Opção 1: Via Duplo Clique (Windows)
Dê um duplo clique no arquivo:
```cmd
iniciar.bat
```

### Opção 2: Pelo PowerShell
```powershell
.\iniciar.ps1
```

### Opção 3: Manualmente via Terminal
```bash
python app.py
```
O sistema abrirá automaticamente o navegador em:
```
http://127.0.0.1:5000
```

---

## 📊 Planilhas de Exemplo e Demonstração

O sistema possui um gerador embutido de planilhas de teste na pasta `exemplos/`:
- `abastecimentos_rs_geral_exemplo.xlsx`: ~200 transações de todo o RS (Porto Alegre, Pelotas, Santa Maria, Caxias do Sul e Bagé).
- `viaturas_46bpm_alvo_exemplo.xlsx`: Viaturas cadastradas da 1ª Cia, 2ª Cia, Pelotão Força Tática e Patrulha Rural do 46º BPM.

Você pode:
- Clicar em **Opções > Carregar Demonstração** na barra superior para popular o banco com 1 clique.
- Baixar os arquivos modelos em **Opções > Baixar Modelo RS / Viaturas**.
- Clicar em **Exportar Excel (.xlsx)** para obter o relatório final formatado com abas de Resumo da Frota, Abastecimentos Detalhados e Alertas.
