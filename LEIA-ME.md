# SKA TECH HUB — Avaliação da 4ª Jornada Tecnológica 2026

Pacote pronto para publicar como **Google Apps Script Web App**, usando a planilha do evento como banco compartilhado.

## Funcionalidades
- Login separado para cada avaliador por nome de acesso.
- 17 projetos lidos automaticamente da aba `Projetos`.
- Avaliadores ativos iniciais: SKA, SENAI, EDUCAÇÃO e CMAIT.
- 6 critérios de 0 a 10, com pesos iguais: Ideia, Desenvolvimento, Aprendizado, Material, Postura e Clareza.
- Média do avaliador = média simples dos 6 critérios.
- Média final = média dos avaliadores ativos.
- Ranking apenas após receber todas as notas previstas.
- Empates mantêm a mesma posição (1º, 1º, 3º).
- Avaliadores não visualizam notas dos outros.
- Painel da organização com progresso e pódio.
- Upload de foto de cada equipe para uma pasta criada no Google Drive.
- Layout responsivo nas cores e com o logo do SKA TECH HUB.

## Instalação
1. Envie `Avaliacao_Projetos_4_Jornada_2026.xlsx` ao Google Drive e abra como Google Planilhas.
2. Na planilha: **Extensões > Apps Script**.
3. Cole `Code.gs` no arquivo de código.
4. Crie um HTML chamado **Index** e cole `Index.html`.
5. Execute uma vez a função `setupSistema()` e autorize o acesso.
6. A função criará as abas `Config`, `Avaliacoes_Web` e `Fotos_Equipes`.
7. Troque os nome de acessos na aba `Config` antes do evento. nome de acessos iniciais: organização `2026`, SKA `1001`, SENAI `1002`, EDUCAÇÃO `1003`, CMAIT `1004`.
8. Apps Script: **Implantar > Nova implantação > Aplicativo da Web**.
9. Executar como: **Você**. Acesso: escolha a opção compatível com a banca; se os avaliadores não usam a mesma conta/domínio, normalmente será necessário permitir acesso pelo link.
10. Compartilhe o único link publicado com todos os avaliadores.

## Teste visual imediato
Abra `demo.html` no Chrome ou Edge. Na demo qualquer nome de acesso funciona e as avaliações ficam apenas no navegador (`localStorage`). O objetivo é testar layout e fluxo. Para vários avaliadores ao mesmo tempo, use a versão Apps Script.

## Observação de segurança
Os nome de acessos são adequados para um evento interno, mas não equivalem a autenticação corporativa. Troque todos antes de publicar e mantenha o nome de acesso de organização restrito.


## Painel da Organização — consulta de resultados

A Organização agora pode consultar as avaliações assim que forem registradas. Em **Resultados por equipe**, o botão **Ver resultados** mostra cada avaliador, as seis notas, a média individual e as observações. A **média parcial** aparece mesmo antes de a equipe receber todas as avaliações; a **média final** e o pódio continuam sendo fechados quando a quantidade esperada de avaliações for atingida.


## Painel da Organização
- Acesso por um único botão **Entrar**, sem nome, PIN ou senha.
- Exibe as notas assim que forem enviadas.
- Permite apagar uma avaliação individual pelo nome do avaliador e projeto.
- Exibe automaticamente os 3 projetos mais bem colocados conforme a média atual; ao completar todas as avaliações, o painel representa o resultado final.
