# Contracheques DP — primeira versão local

Aplicação Windows para separar contracheques pelo nome completo e criar rascunhos ou enviar mensagens individuais pelo Outlook clássico, usando a conta já configurada no computador. Não solicita senha de e-mail.

## Instalação

1. Instale Python 3.11 ou superior no Windows (com o inicializador `py`).
2. Extraia esta pasta em uma pasta interna com acesso restrito ao Departamento Pessoal.
3. Abra o Outlook **clássico**, configurado com a conta remetente da empresa.
4. Execute `iniciar.bat`. A primeira execução precisa de internet para baixar as bibliotecas; o inicializador também verifica dependências nas próximas execuções.

O novo Outlook e o Outlook exclusivamente no navegador não oferecem esta integração. Políticas da empresa podem bloquear automação ou pedir confirmação no Outlook. Nesse caso, o TI precisa avaliar a integração; esta versão não usa Microsoft Graph.

## Arquivos de entrada

Excel `.xlsx`: primeira aba, primeira linha com as colunas `nome` e `email`. Uma linha por funcionário, com o nome completo conforme o contracheque. Somente linhas com e-mail preenchido entram no lote; linhas sem e-mail são ignoradas e relacionadas na conferência. E-mails preenchidos com formato inválido bloqueiam o processamento. A coluna `cpf` pode existir, mas é opcional e não é utilizada na associação. Não use células mescladas ou fórmulas sem valores salvos.

PDF: precisa conter texto selecionável e o nome completo do funcionário em cada página. A comparação ignora acentos, maiúsculas/minúsculas, pontuação e espaços/quebras de linha extras, mas exige todas as palavras do nome, na mesma ordem. Não aproxima nomes, não aceita abreviações e não corrige erros de digitação automaticamente. Todas as páginas associadas ao mesmo nome são reunidas em um anexo.

Páginas sem correspondência com um destinatário do Excel são ignoradas: não geram anexos nem mensagens. Assim, é possível testar com apenas uma pessoa ou enviar apenas aos funcionários que possuem e-mail. A conferência mostra o total de destinatários, as páginas ignoradas e os cadastros sem e-mail; os detalhes ficam em `conferencia.txt` na pasta do lote. Nomes completos repetidos entre os destinatários (inclusive após normalização), páginas ambíguas e destinatários sem contracheque bloqueiam o lote. No modelo de recibo fornecido, a aplicação lê o campo Nome do Funcionário e compara o nome completo exato após normalização. Em outros modelos, busca no texto; se um nome estiver contido em outro e houver mais de uma correspondência, bloqueia a página. Homônimos precisam de outro identificador e de adaptação da aplicação; o CPF da planilha sozinho não resolve se não aparece no PDF.

Não há OCR nesta primeira versão: PDFs escaneados precisam de adaptação. Não se presume que páginas sem identificação pertençam ao funcionário anterior. Revise visualmente todos os anexos e as páginas ignoradas antes de enviar. Uma página ignorada pode corresponder a alguém sem cadastro ou a uma diferença na escrita do nome.

## Uso

1. Selecione o Excel e o PDF e clique em **Processar e conferir arquivos**.
2. Confira os nomes e e-mails. Dê dois cliques nas linhas para abrir os anexos separados.
3. Informe o endereço exato da conta configurada no Outlook e ajuste assunto/mensagem.
4. Para o primeiro teste, use dados fictícios e destinatários internos de teste, e escolha **Criar rascunhos**. Confira os rascunhos no Outlook.
5. Em um lote de produção revisado, escolha **Enviar e-mails agora** e confirme.

Criar rascunhos e enviar são alternativas para o mesmo lote: após qualquer processamento, a interface exige um novo lote. O envio direto cria novas mensagens e não envia rascunhos anteriores. Se criou rascunhos reais, envie-os pelo Outlook; não reprocesse os mesmos arquivos para envio direto, pois isso pode duplicar mensagens.

## Relatórios e cuidados operacionais

Os PDFs separados e `relatorio.csv` ficam em `lotes/`, dentro desta pasta. `envios.jsonl` registra tentativas antes de chamar o Outlook e bloqueia novas mensagens para o mesmo e-mail e o mesmo PDF, inclusive após falha parcial. Os demais destinatários podem continuar: uma tentativa de teste com uma pessoa não bloqueia mais o PDF inteiro. A confirmação informa quantos serão processados e quantos serão ignorados por tentativa anterior. Registros antigos sem identificação do destinatário bloqueiam o PDF por precaução. Não apague esse arquivo para contornar uma falha: verifique Caixa de Saída e Itens Enviados com o TI antes de qualquer reenvio. Alterar o PDF gera outra identificação; esta proteção não detecta duplicidades entre PDFs diferentes nem envios manuais de rascunhos.

“Entregue ao Outlook” significa que o Outlook recebeu a ordem de envio, não confirmação de recebimento pelo funcionário. Outlook offline pode manter mensagens na Caixa de Saída. A aplicação não acompanha devoluções.

Os dados são processados localmente; o Outlook envia os e-mails pelo ambiente da empresa. A pasta contém dados pessoais e salariais: use permissões de acesso do Windows e a política de retenção da empresa. Esta versão não criptografa os PDFs nem implementa autenticação própria.

Antes de usar em produção, valide com um PDF representativo e com o TI. A compatibilidade com os arquivos e a conta corporativa ainda precisa ser conferida no PC de destino.

## Teste com uma pessoa

Selecione o PDF completo e o Excel com apenas o destinatário de teste preenchido. O programa separa somente as páginas dessa pessoa e ignora as demais. Confira a linha na tabela e abra o anexo. Escolha criar rascunho para revisar no Outlook antes de qualquer envio real.

## Atualizar uma instalação existente

Feche a aplicação e substitua somente `app.py`, `requirements.txt`, `iniciar.bat` e `LEIA-ME.md` pelos arquivos da nova versão. Preserve `envios.jsonl` e a pasta `lotes`. Execute `iniciar.bat` novamente. Não apague o histórico de envios para resolver avisos de duplicidade.
