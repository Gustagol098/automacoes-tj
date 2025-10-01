⚖️ Automação Tribunal de Justiça - Gestão de Sistemas e Chamados

Automação desenvolvida para otimizar o processo de renovação de sistemas no Sentinela e a gestão de chamados no GLPI no Tribunal de Justiça.
O objetivo é agilizar tarefas repetitivas, garantindo mais eficiência no setor e reduzindo o tempo gasto com processos manuais, mantendo conformidade com boas práticas de segurança e LGPD.

🚀 Funcionalidades
🔄 Renovação completa

Atualiza automaticamente todos os sistemas principais no Sentinela.
Cria e fecha de forma automática os chamados correspondentes no GLPI.
Ideal para atualizações gerais de sistemas.

🎯 Renovação seletiva

Permite escolher quais sistemas devem ser renovados.
Cria e fecha chamados apenas para os sistemas selecionados.
Ideal para atualizações pontuais ou específicas.

🔒 Conformidade com LGPD e boas práticas

A automação foi desenvolvida seguindo princípios de segurança e privacidade:

Minimização de dados: nenhuma informação pessoal é armazenada ou exposta no código.

Uso restrito: scripts preparados para ambientes institucionais, sem compartilhar dados sensíveis.

Boas práticas de autenticação: credenciais e acessos não estão hardcoded no projeto, devendo ser configurados via variáveis de ambiente ou arquivos externos seguros.

Sigilo de informações: o repositório não contém dados, logs ou prints referentes ao ambiente real do Tribunal de Justiça.

🛠️ Tecnologias utilizadas

sys – manipulação de parâmetros e controle do sistema.
os – gerenciamento de diretórios e variáveis de ambiente.
subprocess – execução de processos externos.
tkinter – interface gráfica simples para seleção de sistemas.
playwright – automação de navegação web (Sentinela e GLPI).
time – controle de espera e sincronização de etapas.

⚠️ Aviso

Este repositório contém apenas a estrutura da automação.
Por motivos de LGPD e segurança institucional, nenhuma credencial, dado sensível ou informação real de sistemas está presente.
