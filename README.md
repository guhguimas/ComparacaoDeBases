📊 Robô de Comparação de Planilhas – Cessão
📌 Sobre o projeto

Este robô foi criado para facilitar a análise da planilha de Cessão, comparando seus dados com duas bases externas.
O objetivo é identificar automaticamente determinados tipos de operação e gerar uma nova planilha já tratada e pronta para uso.

O processo é feito de forma automática, reduzindo erros manuais e economizando tempo operacional.

🎯 O que o robô faz

O robô realiza as seguintes ações:

Lê a planilha principal de Cessão

Lê duas bases externas em formato CSV

Remove os traços da coluna de identificação (CCB)

Compara os dados entre as planilhas

Identifica o tipo de operação correspondente

Cria uma nova coluna com essa informação

Filtra apenas os casos relevantes

Gera uma nova planilha final já pronta para análise

📂 Arquivos utilizados

O robô utiliza:

Planilha Cessão (formato Excel)

Base 1 (formato CSV)

Base 2 (formato CSV)

O resultado será:

Uma nova planilha Excel contendo apenas os registros filtrados

🧠 Regras aplicadas automaticamente

O robô irá selecionar apenas os casos onde a operação for:

DIG

GDC

SEM CESSAO

CAPITAL

Não encontrada (#N/D)

Vazia

Esses registros serão separados em uma nova planilha.

🧾 Resultado final

A planilha gerada:

Mantém o mesmo visual da planilha original

Contém apenas os registros necessários

Inclui uma nova coluna chamada:

OPERACAO FRONT

▶️ Como utilizar

Abrir o robô

Selecionar:

A planilha Cessão

A Base 1

A Base 2

A pasta onde o arquivo será salvo

Clicar em Iniciar

Aguardar o processamento

O arquivo final será gerado automaticamente

🛑 Cancelamento

Caso necessário, o processo pode ser cancelado durante a execução.

📊 Resumo automático

Ao final da execução, o robô apresenta um resumo com a quantidade de registros encontrados por tipo de operação.

💡 Benefícios

Redução de erros humanos

Processamento muito mais rápido

Padronização das análises

Facilidade operacional

Planilha final pronta para uso

Economia de tempo

🔄 Evolução do projeto

Este robô foi desenvolvido para evoluir conforme novas necessidades operacionais surgirem.

Novas regras, filtros e integrações poderão ser adicionadas ao longo do tempo.