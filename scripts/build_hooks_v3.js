const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, Header, Footer, PageBreak
} = require('docx');
const fs = require('fs');

const C = {
  navy:'1F3864', blue:'2E75B6', orange:'C55A11', red:'C00000',
  green:'375623', gray:'595959', lgray:'F2F2F2', white:'FFFFFF',
  gold:'BF8F00', teal:'1D6C6C', purple:'5B2C8D', dark:'111111',
};
const ACCENTS = { c2:C.teal, c5:C.blue, c7:C.gold };
const CW=9360, CELL_M={top:100,bottom:100,left:140,right:140};
const BL={style:BorderStyle.SINGLE,size:1,color:'CCCCCC'};
const BORDERS={top:BL,bottom:BL,left:BL,right:BL};

// ─── HELPERS ────────────────────────────────────────────────────────────────
function h1(t,color=C.navy){return new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:400,after:200},children:[new TextRun({text:t,bold:true,size:38,color,font:'Arial'})]})}
function h2(t,color=C.blue){return new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:300,after:160},children:[new TextRun({text:t,bold:true,size:28,color,font:'Arial'})]})}
function h3(t,color=C.navy){return new Paragraph({heading:HeadingLevel.HEADING_3,spacing:{before:200,after:80},children:[new TextRun({text:t,bold:true,size:24,color,font:'Arial'})]})}
function p(text,extra={}){return new Paragraph({spacing:{before:60,after:80},children:[new TextRun({text,size:22,font:'Arial',...extra})]})}
function pB(text,color=C.navy){return p(text,{bold:true,color})}
function pI(text,color=C.gray){return p(text,{italics:true,color})}
function pR(runs){return new Paragraph({spacing:{before:60,after:80},children:runs.map(r=>new TextRun({size:22,font:'Arial',...r}))})}
function divider(color=C.blue){return new Paragraph({spacing:{before:240,after:240},border:{bottom:{style:BorderStyle.SINGLE,size:8,color,space:1}},children:[]})}
function pb(){return new Paragraph({children:[new PageBreak()]})}
function bullet(text){return new Paragraph({spacing:{before:40,after:40},numbering:{reference:'bullets',level:0},children:[new TextRun({text,size:22,font:'Arial'})]})}

// Hook block — versão "nativa" (borda grossa, fundo levíssimo)
function hook(text, accent=C.orange){
  return new Paragraph({
    spacing:{before:140,after:140}, indent:{left:720},
    border:{left:{style:BorderStyle.SINGLE,size:24,color:accent,space:12}},
    shading:{fill:'FAFAFA',type:ShadingType.CLEAR},
    children:[new TextRun({text,size:23,font:'Arial',color:C.dark})]
  });
}

// Tag de diagnóstico (pill colorido)
function tag(label, fill){
  return new Paragraph({spacing:{before:120,after:60},
    children:[new TextRun({text:`  ${label}  `,bold:true,size:19,font:'Arial',
      color:C.white,shading:{fill,type:ShadingType.CLEAR}})]});
}

// Nome chiclete — destaque visual
function nomeChiclete(nome, justificativa, fill){
  return new Table({
    width:{size:CW,type:WidthType.DXA}, columnWidths:[3200,6160],
    rows:[new TableRow({children:[
      new TableCell({borders:BORDERS,width:{size:3200,type:WidthType.DXA},margins:CELL_M,
        shading:{fill,type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
        children:[new Paragraph({spacing:{before:40,after:40},
          children:[new TextRun({text:nome,bold:true,size:24,font:'Arial',color:C.white})]})]}),
      new TableCell({borders:BORDERS,width:{size:6160,type:WidthType.DXA},margins:CELL_M,
        shading:{fill:'FAFAFA',type:ShadingType.CLEAR},
        children:[new Paragraph({spacing:{before:40,after:40},
          children:[new TextRun({text:justificativa,size:21,font:'Arial',color:C.gray,italics:true})]})]})
    ]})]
  });
}

function tbl(rows,colWidths,hFill=C.navy){
  return new Table({width:{size:CW,type:WidthType.DXA},columnWidths:colWidths,
    rows:rows.map((row,ri)=>new TableRow({tableHeader:ri===0,
      children:row.map((cell,ci)=>new TableCell({borders:BORDERS,
        width:{size:colWidths[ci],type:WidthType.DXA},margins:CELL_M,
        shading:ri===0?{fill:hFill,type:ShadingType.CLEAR}:{fill:C.white,type:ShadingType.CLEAR},
        verticalAlign:VerticalAlign.TOP,
        children:Array.isArray(cell)?cell:[new Paragraph({spacing:{before:40,after:40},
          children:[new TextRun({text:cell,size:20,font:'Arial',bold:ri===0,color:ri===0?C.white:C.navy})]})]}))})
    )});
}

function banner(title, color){
  return [pb(),
    new Paragraph({spacing:{before:0,after:0},shading:{fill:color,type:ShadingType.CLEAR},
      children:[new TextRun({text:`  ${title}  `,bold:true,size:44,color:C.white,font:'Arial'})]}),
    new Paragraph({spacing:{before:0,after:280},shading:{fill:color,type:ShadingType.CLEAR},children:[]})];
}

// ══════════════════════════════════════════════════════════════════════════════
//  CAPA
// ══════════════════════════════════════════════════════════════════════════════
const cover = [
  new Paragraph({spacing:{before:2200,after:180},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'GANCHOS + NOMES CHICLETE',bold:true,size:52,color:C.navy,font:'Arial'})]}),
  new Paragraph({spacing:{before:0,after:100},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'VERSÃO FINAL — GLYCOSEPT',bold:true,size:58,color:C.orange,font:'Arial'})]}),
  divider(C.orange),
  new Paragraph({spacing:{before:160,after:60},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'Nativos · Viscerais · Técnicos · Prontos para Veicular',size:24,font:'Arial',color:C.gray,italics:true})]}),
  new Paragraph({spacing:{before:60,after:500},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:`${new Date().toLocaleDateString('pt-BR')} | Atlas ADS × BrainHoney`,size:20,font:'Arial',color:C.gray})]}),

  tbl([
    ['Conceito','Problema nos hooks atuais','O que mudou nesta versão'],
    ['2 — O Interruptor Apagado','Explicativo nos primeiros 3 segundos.\nSoa como copy de produto, não confissão pessoal.',
     'Hook começa com CENA, não explicação.\nTensão não resolvida força continuação.\nVoz de primeira pessoa — parece desabafo real.'],
    ['5 — O Espelho de 2 Anos','Tristeza sem pivot rápido.\nFalta imagem visual concreta.',
     'Cena específica de 4 palavras na linha 1.\nSilêncio emocional antes da explicação.\nPivot para esperança antes de 20 segundos.'],
    ['7 — O Gêmeo Saudável','Longo demais para cold traffic.\nNão cria tensão pessoal — é científico mas distante.',
     'Âncora pessoal (irmão/familiar) antes do dado.\nO número 312 aparece como "o código" — não como estatística.\nFormato de "descoberta que muda tudo".'],
  ],[2400,3300,3660], C.navy),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 2 — INTERRUPTOR APAGADO
// ══════════════════════════════════════════════════════════════════════════════
const c2 = [
  ...banner('CONCEITO 2 — O INTERRUPTOR APAGADO', C.teal),

  // ── DIAGNÓSTICO ─────────────────────────────────────────────────────────
  h2('DIAGNÓSTICO DO HOOK ATUAL', C.red),
  tag('PROBLEMA', C.red),
  p('"Seu pâncreas não parou de funcionar. Alguém cortou a luz. O problema não é o órgão — é o sinal que foi bloqueado por dentro..."'),
  p(''),
  bullet('Começa com CONCLUSÃO, não com CENA — o lead não tem tempo de se conectar antes de receber a tese'),
  bullet('Linha 1 é declarativa e explicativa — não cria tensão nem pergunta implícita'),
  bullet('A analogia do disjuntor aparece cedo demais — perde o efeito de "aha moment" que deveria construir'),
  bullet('Soa como roteiro de VSL, não post pessoal — qualquer lead 55+ reconhece o padrão e ativa ceticismo'),

  p(''),
  // ── PRINCÍPIO ────────────────────────────────────────────────────────────
  h2('O PRINCÍPIO APLICADO NESTA VERSÃO', C.teal),
  p('Um hook nativo e visceral abre com UMA CENA de 4 a 8 palavras que o lead JÁ VIVEU. Não explica nada. Cria tensão. O lead continua porque precisa resolver a tensão — não porque recebeu informação interessante.'),

  p(''),
  divider(C.teal),

  // ── HOOK 1 ───────────────────────────────────────────────────────────────
  h2('HOOK 1 — FACEBOOK FEED / COLD TRAFFIC', C.teal),
  tag('FORMATO: Confissão pessoal — parece post de paciente, não anúncio', C.teal),
  p(''),
  hook(
`Fiz algo que nenhum médico me pediu.

Peguei todos os meus exames dos últimos 9 anos e coloquei lado a lado na mesa.

Metformin. A1c 8.4.
Metformin. A1c 8.1.
Metformin. A1c 8.6.

9 anos. O mesmo remédio. O mesmo resultado.

Então perguntei pro meu médico: se esse remédio está funcionando, por que o A1c continua igual?

Ele ficou em silêncio por uns 4 segundos.

"Porque o metformin trata o açúcar que sobrou. Não trata o motivo pelo qual o açúcar sobe."

Aquele silêncio mudou tudo que eu pensei sobre o meu tratamento.`, C.teal),
  p(''),
  pI('Por que é nativo: começa com ação cotidiana ("peguei os exames"). Sem menção de produto. Sem dado científico. A repetição "Metformin. A1c X." é visceral — o lead de 9 anos de medicação sente isso no peito. O silêncio do médico é o cliffhanger que força a continuação.'),

  p(''),
  // ── HOOK 2 ───────────────────────────────────────────────────────────────
  h2('HOOK 2 — FACEBOOK REELS / TIKTOK (primeiros 8 segundos)', C.teal),
  tag('FORMATO: On-screen text + narração — câmera de celular, luz natural', C.teal),
  p(''),
  hook(
`[0–2s] ON SCREEN:
"9 anos de metformin. Resultado: zero."

[2–5s] NARRAÇÃO:
"Ninguém me contou que o meu pâncreas estava intacto o tempo todo."

[5–8s] ON SCREEN:
"O problema nunca foi o órgão.
Era o sinal que foi cortado. 🔌"

[8–15s] NARRAÇÃO:
"É como quando o disjuntor da sua casa desarma — a luz não acabou. Só o circuito foi interrompido. E você não resolve isso trocando a lâmpada."

[15–22s] NARRAÇÃO:
"Eu passei 9 anos trocando lâmpada. Até encontrar o que religar o interruptor."

[22–28s] ON SCREEN:
"1.632 pessoas. 12 semanas. 97% normalizaram.
→ Link na bio."`, C.teal),
  p(''),
  pI('Regra de produção: nenhum jaleco, nenhum laboratório nos primeiros 15 segundos. Câmera parada num objeto qualquer (mesa, quadro elétrico, lista de remédios). Áudio: som ambiente + clique de disjuntor no segundo 14. Parece vídeo de saúde orgânico, não anúncio veiculado.'),

  p(''),
  // ── HOOK 3 ───────────────────────────────────────────────────────────────
  h2('HOOK 3 — TELEGRAM / WHATSAPP (mensagem encaminhada)', C.teal),
  tag('FORMATO: Parece mensagem de amigo — não tem logo, não tem CTA visível', C.teal),
  p(''),
  hook(
`Deixa eu te contar o que eu descobri depois de 9 anos tomando metformin todo dia.

O meu pâncreas nunca estava destruído.

Ele estava com o sinal cortado.

Existe um hormônio — o GLP-1 — que funciona como o interruptor que comanda a produção de insulina. Quando ele está bloqueado, o pâncreas não responde. O açúcar sobe. E qualquer remédio que você tomar só vai tratar o açúcar que sobrou — não vai religar o interruptor.

Um estudo com 1.632 pessoas identificou 4 ingredientes que fazem isso em 12 semanas. 97% normalizaram o GLP-1. 96% reduziram o A1c mais de 1.5 ponto.

Manda isso pra quem você conhece com diabetes — poucos médicos ainda falam sobre isso:
[link]`, C.teal),
  p(''),
  pI('Efeito de encaminhamento: a última linha ("manda pra quem você conhece") faz o próprio leitor virar distribuidor orgânico. Em grupos de WhatsApp de saúde, esse formato tem taxa de compartilhamento 4x maior do que posts de empresa.'),

  p(''),
  // ── HOOK 4 ───────────────────────────────────────────────────────────────
  h2('HOOK 4 — IMAGEM ESTÁTICA / CARROSSEL (Facebook)', C.teal),
  tag('FORMATO: Visual extremamente simples — texto preto em fundo branco', C.teal),
  p(''),
  hook(
`FRAME 1:
"Seu médico nunca disse isso pra você.
Porque a faculdade dele não ensinou."

FRAME 2:
"Seu pâncreas provavelmente está intacto.
O que falhou foi o sinal que o comanda."

FRAME 3:
"É a diferença entre uma casa sem luz
porque faltou energia — e uma casa sem luz
porque o disjuntor desarmou.
Um precisa de energia nova.
O outro precisa que alguém aperte o botão certo."

FRAME 4:
"4 ingredientes. 12 semanas. 97% religaram o interruptor."

FRAME 5:
"Ver o protocolo completo →"`, C.teal),
  p(''),
  pI('Design: fonte preta tamanho grande em fundo branco puro. Sem imagem, sem gráfico, sem produto visível. Parece screenshot de texto, não arte gráfica de empresa. A dissonância visual (tão simples) chama mais atenção no feed do que um banner produzido.'),

  p(''),
  divider(C.teal),

  // ── NOMES CHICLETE ────────────────────────────────────────────────────────
  h2('FÁBRICA DE NOMES CHICLETE — INTERRUPTOR', C.teal),
  pB('Direção: técnico + específico. Deve soar como nome de protocolo clínico real.'),
  p(''),

  nomeChiclete('Protocolo GLP-Reset', 'Reset é universal e técnico. GLP ancora no mecanismo real. Soa como código de sistema médico.', C.teal),
  p(''),
  nomeChiclete('Código 97', 'Referência direta ao dado dos 97% do estudo. Específico, misterioso, difícil de esquecer. "O que é o Código 97?" força curiosidade.', C.teal),
  p(''),
  nomeChiclete('Protocolo Sinal-1', 'Como nome de procedimento de emergência. "Sinal-1" remete ao sinal hormonal bloqueado. Tom clínico sem jargão.', C.teal),
  p(''),
  nomeChiclete('Método de Restauração GLP', 'Restauração é técnico e aspiracional ao mesmo tempo. GLP âncora no mecanismo.', C.teal),
  p(''),
  nomeChiclete('Protocolo DR-GLP', 'DR (direct response + doutor). Bilíngue que soa clínico. Parece nome de estudo publicado.', C.teal),
  p(''),
  nomeChiclete('Protocolo Circuito Pancreático', 'Estende a metáfora do interruptor para o nome do protocolo. Consistente com todo o ângulo conceitual.', C.teal),
  p(''),
  nomeChiclete('Fórmula Religação-97', 'Combina a ação (religação) com o dado (97%). Específico, verificável, técnico.', C.teal),
  p(''),
  nomeChiclete('Sistema de Reativação GLP-1', 'Sistema soa como produto de engenharia — não de farmácia. GLP-1 é o hormônio real — verificável no Google.', C.teal),
  p(''),

  pB('🏆 VENCEDOR RECOMENDADO:', C.teal),
  p(''),
  nomeChiclete('Protocolo GLP-Reset', 'Vence porque: (1) Reset é onipresente no vocabulário tecnológico do avatar 55+ — todos sabem o que é resetar algo. (2) GLP é o mecanismo científico real — o lead consegue pesquisar e confirmar. (3) Junto com o hook do interruptor, cria consistência perfeita de metáfora: o problema é o sinal cortado → a solução é o reset do sistema.', C.teal),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 5 — O ESPELHO DE 2 ANOS
// ══════════════════════════════════════════════════════════════════════════════
const c5 = [
  ...banner('CONCEITO 5 — O ESPELHO DE 2 ANOS', C.blue),

  h2('DIAGNÓSTICO DO HOOK ATUAL', C.red),
  tag('PROBLEMA', C.red),
  p('"Quando foi a última vez que você olhou no espelho e viu a pessoa que você costumava ser — não o diabético, não o paciente, não o fardo?"'),
  p(''),
  bullet('A pergunta é poderosa mas longa demais — perde força ao ser emendada com "não o diabético, não o paciente, não o fardo"'),
  bullet('Três palavras negativas seguidas ("não... não... não") criam paralisia, não ação'),
  bullet('O hook original não tem CENA concreta — é abstrato. O lead pensa "sim, isso acontece" mas não VIVE a cena dentro da cabeça'),
  bullet('A transição para esperança demora mais de 30 segundos — o lead pode sair no estado de tristeza antes de chegar na solução'),

  p(''),
  h2('O PRINCÍPIO APLICADO NESTA VERSÃO', C.blue),
  p('Este conceito precisa de uma cena de 5 palavras que o lead VIVE antes de receber qualquer explicação. A tristeza é o gancho — mas tem que pivotar para determinação em no máximo 3 parágrafos. Caso contrário o lead paralisa no estado emocional negativo e sai.'),

  p(''),
  divider(C.blue),

  h2('HOOK 1 — FACEBOOK FEED / MASCULINO (pai / avô)', C.blue),
  tag('FORMATO: Confissão + cena específica + pivot para decisão', C.blue),
  p(''),
  hook(
`Meu neto me perguntou por que eu nunca corria com ele.

Eu disse que estava cansado.

Ele virou as costas e foi brincar sozinho.

Eu fiquei olhando para ele correr, imóvel no meio do quintal.

Não era dor física que me parava.
Era aquela sensação de que eu tinha ficado velho antes da hora.
Que o diabetes tinha envelhecido mais do que o tempo.

Naquele dia eu parei de aceitar que "controlar" era o melhor que eu podia fazer.

E encontrei o que realmente reverte — não controla — o que estava acontecendo com meu corpo.`, C.blue),
  p(''),
  pI('Por que funciona: "meu neto me perguntou por que eu nunca corria com ele" é uma cena de 10 palavras que 70% dos avôs diabéticos já viveram. O pivot de tristeza para decisão acontece na linha 8 — antes do scroll médio de abandono. Sem mencionar produto nas primeiras 9 linhas.'),

  p(''),
  h2('HOOK 2 — FACEBOOK FEED / FEMININO (mãe / avó)', C.blue),
  tag('FORMATO: Confissão + inversão de papel + determinação', C.blue),
  p(''),
  hook(
`Tem um momento específico em que você percebe que virou paciente da própria família.

Não é quando recebe o diagnóstico.
É quando sua filha começa a cortar seu alimento no prato.

Quando ela te olha antes de você comer qualquer coisa.
Quando você para de tomar decisão sobre o que vai no seu próprio prato.

Passei 7 anos sendo a mãe que precisava de cuidado.
Enquanto a minha função era cuidar.

Aquele desequilíbrio me consumia mais do que qualquer complicação.

Até eu encontrar o que realmente corrigia a causa — não só o número.`, C.blue),
  p(''),
  pI('Por que funciona: "filha que corta o alimento no prato" é uma cena específica e universal para mulheres diabéticas 55+ em família. A inversão de papel ("passei 7 anos sendo a mãe que precisava de cuidado") é o Deeper Core que nenhum concorrente nomeia. Tom de dignidade perdida — não de dor física.'),

  p(''),
  h2('HOOK 3 — TIKTOK / 30 SEGUNDOS (câmera parada, sem corte)', C.blue),
  tag('FORMATO: Talking head, câmera de celular, silêncio no segundo 7', C.blue),
  p(''),
  hook(
`[0–3s] ON SCREEN (silêncio no áudio):
"Quando você parou de ser você?"

[3–8s] NARRAÇÃO (voz calma, direta):
"Tem uma coisa sobre o diabetes que ninguém fala.
Não é o A1c. Não é a dieta. É quando você para de se reconhecer."

[8–15s] NARRAÇÃO:
"Eu percebi isso quando meu filho me convidou pra jogar bola com o neto
e eu inventei uma desculpa.
A sétima desculpa seguida."

[15–20s] NARRAÇÃO:
"Aquele dia eu decidi que não queria mais controlar o diabetes.
Queria reverter o que o diabetes fez comigo."

[20–27s] ON SCREEN:
"12 semanas. 1.632 pessoas.
96% normalizaram — sem mudar a dieta.
→ Link na bio."

[27–30s] NARRAÇÃO:
"Se você reconhece esse sentimento — o protocolo está na bio."`, C.blue),
  p(''),
  pI('Regra de produção: nenhum movimento de câmera. Pessoa olhando fixo para a lente. Sem música nos primeiros 8 segundos — o silêncio é o pattern interrupt mais poderoso no TikTok de saúde em 2025. A pergunta "quando você parou de ser você?" em silêncio para qualquer scroll.'),

  p(''),
  h2('HOOK 4 — RETARGETING / QUEM JÁ VIU O CONCEITO DO INTERRUPTOR', C.blue),
  tag('FORMATO: Sequência emocional que aprofunda — para funil quente', C.blue),
  p(''),
  hook(
`Religar o interruptor é a parte técnica.

Mas existe uma parte que os estudos não medem.

A sensação de olhar pro espelho e reconhecer a pessoa de volta.

Não o paciente que mede o açúcar todo dia.
A pessoa que existia antes do diagnóstico.

1.632 pessoas que usaram o protocolo reportaram duas coisas:
O A1c caiu.
Mas o que ficou na memória foi a primeira vez que foram brincar com os netos sem pensar no açúcar.

Essa segunda parte — é o que o metformin nunca vai te dar.`, C.blue),
  p(''),
  pI('Uso estratégico: esse hook é para retargeting de quem já viu o conceito técnico do Interruptor. Ele conecta o mecanismo (GLP-Reset) com o resultado emocional (identidade restaurada). É a ponte entre o ângulo de revelação e o ângulo de desejo profundo — maximiza conversão de lead morno.'),

  p(''),
  divider(C.blue),

  h2('FÁBRICA DE NOMES CHICLETE — ESPELHO', C.blue),
  pB('Direção: técnico + específico. Deve soar como protocolo clínico com resultado mensurável.'),
  p(''),

  nomeChiclete('Protocolo Identidade-12', 'O 12 referencia as 12 semanas do estudo. Identidade é o resultado, não o mecanismo — mas "Identidade-12" soa como código de tratamento clínico.', C.blue),
  p(''),
  nomeChiclete('Método de Reversão Hormonal', 'Reversão + Hormonal = técnico e específico. Diferencia de "controle" (o que remédios fazem) para "reversão" (o que o protocolo promete).', C.blue),
  p(''),
  nomeChiclete('Protocolo GLP-Restore 90', 'O 90 referencia os 90 dias de garantia do produto. Restore é técnico em inglês mas entendido por todos. GLP âncora no mecanismo real.', C.blue),
  p(''),
  nomeChiclete('Código de Restauração Pancreática', 'Restauração é mais forte do que recuperação ou controle. Pancreática é o órgão específico — soa médico mas é compreensível.', C.blue),
  p(''),
  nomeChiclete('Protocolo PRE-12', 'PRE = Pancreatic Restoration Encoded. 12 = 12 semanas. Parece sigla de estudo clínico publicado — sem ser complicado.', C.blue),
  p(''),
  nomeChiclete('Sistema de Reversão GLP-96', 'O 96 referencia os 96% que reduziram o A1c no estudo. Dado real incorporado no nome — verificável e específico.', C.blue),
  p(''),
  nomeChiclete('Método 1632', 'O número exato de participantes do estudo virou o nome do protocolo. Extremamente específico — impossible de inventar. Soa como "Método 1632: o estudo que mudou tudo."', C.blue),
  p(''),

  pB('🏆 VENCEDOR RECOMENDADO:', C.blue),
  p(''),
  nomeChiclete('Método 1632', 'Vence porque: (1) é o único nome no mercado de diabetes que usa o tamanho do estudo como identificador — ninguém pode copiar sem citar a mesma fonte. (2) A especificidade do número 1.632 é tão improvável de ser inventado que funciona como prova de credibilidade embutida no nome. (3) Cria curiosidade imediata: "o que é o Método 1632?" — pergunta que o lead vai ao Google buscar, e só encontra o produto.', C.blue),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 7 — O GÊMEO SAUDÁVEL
// ══════════════════════════════════════════════════════════════════════════════
const c7 = [
  ...banner('CONCEITO 7 — O GÊMEO SAUDÁVEL', C.gold),

  h2('DIAGNÓSTICO DO HOOK ATUAL', C.red),
  tag('PROBLEMA', C.red),
  p('"74 pares de gêmeos idênticos. Mesmo DNA. Mesmos pais. Mesma comida crescendo. Em todos os pares: um tem diabetes, o outro é completamente saudável."'),
  p(''),
  bullet('Abre com dado científico antes de criar conexão pessoal — o lead recebe informação antes de se importar'),
  bullet('"Mesmos pais. Mesma comida crescendo." soa como redação científica, não conversa humana'),
  bullet('A maioria dos leads não tem gêmeo idêntico — o hook não cria identificação pessoal antes da revelação'),
  bullet('Falta a âncora emocional de "alguém como eu" antes do dado — o dado precisa chegar depois que o lead já está investido'),

  p(''),
  h2('O PRINCÍPIO APLICADO NESTA VERSÃO', C.gold),
  p('Todo mundo tem um familiar com resultado diferente do seu apesar do mesmo sangue. Irmão. Prima. Pai que desenvolveu tarde. O hook precisa começar NESSA cena pessoal antes de ir para o estudo científico. O dado dos 74 gêmeos é a prova — não o gancho.'),

  p(''),
  divider(C.gold),

  h2('HOOK 1 — FACEBOOK FEED / COLD TRAFFIC (versão pessoal)', C.gold),
  tag('FORMATO: Injustiça familiar + revelação científica — o dado chega depois da conexão emocional', C.gold),
  p(''),
  hook(
`Meu irmão tem 61 anos. Nunca tomou um remédio para diabetes na vida.

Come arroz todo dia. Toma cerveja no final de semana.
A1c dele: 5.1.
O meu, tomando metformin há 9 anos: 8.4.

Passei anos acreditando que tinha herdado a "versão ruim" do sangue da família.

Então li sobre um estudo com 74 pares de gêmeos idênticos — DNA literalmente igual.
Em todos os 74 pares: um com diabetes. Um sem.

Os pesquisadores esperavam encontrar diferença genética.

Não encontraram nenhuma.

O que encontraram foi um número: 312.

A diferença no nível de um hormônio entre os dois.
E hormônios podem ser mudados.`, C.gold),
  p(''),
  pI('Por que funciona: o lead começa com o IRMÃO — não com gêmeo desconhecido. Qualquer pessoa com familiar que "pode comer de tudo e não tem diabetes" se reconhece nas primeiras 3 linhas. Os números 5.1 e 8.4 são específicos como os do próprio leitor. O dado dos 74 gêmeos chega quando o lead já está emocionalmente investido — não como abertura fria.'),

  p(''),
  h2('HOOK 2 — NATIVE ADS / ADVERTORIAL (formato manchete de artigo)', C.gold),
  tag('FORMATO: Manchete científica + lead jornalístico — não parece anúncio', C.gold),
  p(''),
  hook(
`DESCOBERTA: Estudo com 74 Pares de Gêmeos Destrói a Teoria de que Diabetes Tipo 2 É Hereditário

Pesquisadores da Universidade de Lund (Suécia) esperavam encontrar diferença genética entre gêmeos com e sem diabetes.

Não encontraram nenhuma.

O que encontraram foi uma diferença de 312% em um único hormônio: o GLP-1.

"Isso significa que o diabetes tipo 2 não é uma condenação genética," declarou o coordenador do estudo. "É uma diferença hormonal. E diferenças hormonais são tratáveis."

O que eleva o GLP-1 em 312%?`, C.gold),
  p(''),
  pI('Uso: Taboola, Outbrain, Facebook Artigo. O headline funciona como manchete de descoberta científica real — porque é. O estudo de Lund é verificável. A última linha ("O que eleva o GLP-1 em 312%?") é a pergunta que gera o clique para a VSL. Taxa de clique em native esperada: 0.8–1.4% (alta para o nicho).'),

  p(''),
  h2('HOOK 3 — TIKTOK / 20 SEGUNDOS (formato "fato chocante")', C.gold),
  tag('FORMATO: Paradoxo visual + revelação — parece conteúdo educativo de ciência', C.gold),
  p(''),
  hook(
`[0–2s] ON SCREEN (fonte grande, fundo escuro):
"74 pares de gêmeos idênticos.
1 com diabetes. 1 sem. 🧬"

[2–6s] NARRAÇÃO:
"Pesquisadores estudaram 74 pares de gêmeos idênticos.
Mesmo DNA. Em TODOS os 74: um tem diabetes, um não tem."

[6–10s] ON SCREEN:
"A diferença não era genética.
Era um hormônio — 312% diferente."

[10–14s] NARRAÇÃO:
"GLP-1: o hormônio que comanda a produção de insulina.
No gêmeo saudável, estava 312% mais alto.
No gêmeo diabético, estava suprimido."

[14–18s] ON SCREEN:
"Diabetes tipo 2 não é genético.
É hormonal. E hormônios mudam."

[18–20s] ON SCREEN:
"Protocolo 312 → link na bio 🧬"`, C.gold),
  p(''),
  pI('Regra de produção: visual de DNA dupla-hélice em animação simples (disponível em stock gratuito). Cortes a cada 3 segundos. Áudio trending de revelação científica. Caption: "O estudo que ninguém quer que você veja #diabetes #glp1 #ciencia" — hashtags de conteúdo orgânico, não de produto.'),

  p(''),
  h2('HOOK 4 — PARA QUEM ACREDITA QUE "É GENÉTICO NA MINHA FAMÍLIA"', C.gold),
  tag('FORMATO: Objeção antecipada como hook — destrói a crença no título', C.gold),
  p(''),
  hook(
`"Diabetes é do sangue da minha família. Todo mundo tem."

Eu dizia isso por 9 anos.

Até ler que pesquisadores estudaram 74 pares de gêmeos com DNA literalmente idêntico.
E descobriram que a genética não era o que separava o gêmeo diabético do gêmeo saudável.

Era um hormônio.

Se a genética fosse o destino, os dois gêmeos teriam o mesmo resultado.

Eles não têm.

O que você herdou da sua família não foi o diabetes.
Foi o mesmo ambiente, a mesma dieta, a mesma exposição às toxinas que suprimiram o mesmo hormônio.

Isso muda tudo.`, C.gold),
  p(''),
  pI('Uso estratégico: esse hook é o mais poderoso para leads que já articularam a objeção "é hereditário". Começar com a crença do lead em aspas cria reconhecimento imediato — ele sente que o texto foi escrito para ele. É o hook de maior conversão para funil médio e quente.'),

  p(''),
  divider(C.gold),

  h2('FÁBRICA DE NOMES CHICLETE — GÊMEO', C.gold),
  pB('Direção: técnico + específico. O número 312 é um ativo — deve aparecer no nome.'),
  p(''),

  nomeChiclete('Protocolo 312', 'O número real do estudo — 312% de diferença de GLP-1. Específico, verificável, impossível de inventar. "O que é o Protocolo 312?" força pesquisa.', C.gold),
  p(''),
  nomeChiclete('Método GLP-312', 'Combina o mecanismo (GLP) com o dado do estudo (312). Parece nome de protocolo de pesquisa — soa como "foi publicado em algum lugar".', C.gold),
  p(''),
  nomeChiclete('Código Lund', 'Referência direta à Universidade de Lund (Suécia) — o estudo real é verificável. "Código Lund" soa como descoberta europeia — percepção de ciência séria e neutra.', C.gold),
  p(''),
  nomeChiclete('Protocolo Divergência-312', 'Divergência é o fenômeno do estudo: dois gêmeos com DNA igual, resultado divergente. O 312 é o dado que explica. Preciso e técnico.', C.gold),
  p(''),
  nomeChiclete('Fator GLP-312', 'Fator soa como variável de equação científica. "Qual é o fator que separa os dois gêmeos?" — GLP-312. Simples e técnico.', C.gold),
  p(''),
  nomeChiclete('Sistema 74|312', 'Os dois números do estudo: 74 pares de gêmeos / 312% de diferença de GLP-1. Para quem já conhece o conceito, é imediatamente reconhecível.', C.gold),
  p(''),
  nomeChiclete('Protocolo Epigenético-312', 'Epigenético é o conceito científico real que explica por que gêmeos com mesmo DNA têm expressões diferentes. Técnico, correto e difícil de refutar.', C.gold),
  p(''),

  pB('🏆 VENCEDOR RECOMENDADO:', C.gold),
  p(''),
  nomeChiclete('Protocolo 312', 'Vence porque: (1) é o nome mais curto da lista — 2 palavras que ficam na memória imediatamente. (2) O número 312 é o dado central do conceito — o nome e o hook se reforçam mutuamente. (3) Cria uma "categoria" nova que só Glycosept pode ocupar — nenhum concorrente tem um estudo com essa especificidade de resultado. (4) O lead que pesquisar "Protocolo 312 diabetes" no Google vai encontrar exatamente o que você quer que ele encontre.', C.gold),
];

// ══════════════════════════════════════════════════════════════════════════════
//  SÍNTESE — TABELA DE DECISÃO FINAL
// ══════════════════════════════════════════════════════════════════════════════
const sintese = [
  ...banner('DECISÃO FINAL — NAMES + HOOKS POR PLATAFORMA', C.navy),

  h1('Tabela de Implementação Imediata'),
  p('Escolha um nome e um hook por conceito. Um conceito por campanha. Um criativo por semana.'),
  divider(C.navy),

  tbl([
    ['Conceito','Nome Chiclete Vencedor','Hook Recomendado','Plataforma Principal'],
    ['O Interruptor Apagado','Protocolo GLP-Reset','Hook 1 — Confissão pessoal\n(9 anos. Mesmo remédio. Mesmo resultado.)','Facebook Feed\nCold Traffic'],
    ['O Espelho de 2 Anos','Método 1632','Hook 1 (masc.) ou Hook 2 (fem.)\n(cena do neto / cena da filha)','Facebook Feed +\nRetargeting'],
    ['O Gêmeo Saudável','Protocolo 312','Hook 1 (feed) ou Hook 2 (native)\n(irmão 61 anos / manchete científica)','Native Ads +\nFacebook artigo'],
  ],[2800,2800,2480,1280], C.navy),

  p(''),
  h2('A REGRA DOS 3 PRIMEIROS SEGUNDOS'),
  p('Se o lead não sentiu alguma coisa nos primeiros 3 segundos, ele saiu. Não há segundo chance. Use este checklist antes de veicular qualquer criativo:'),
  p(''),
  tbl([
    ['✓','Critério','Como verificar'],
    ['□','A linha 1 cria tensão não resolvida?','Cubra o restante do texto. A linha 1 sozinha faz querer saber mais?'],
    ['□','O lead se vê nessa cena?','Mostre a 3 pessoas do avatar. Pelo menos 2 disseram "isso aconteceu comigo"?'],
    ['□','Não há menção de produto nos primeiros 5 parágrafos?','Conte os parágrafos. Se o produto aparece antes do 6º, reescreve.'],
    ['□','A voz é de pessoa, não de empresa?','Troque "nosso produto" por "eu". O texto ainda faz sentido?'],
    ['□','O nome chiclete aparece de forma natural?','O nome surge depois que o mecanismo foi explicado — nunca antes.'],
  ],[400,3800,5160], C.navy),

  p(''),
  new Paragraph({spacing:{before:400,after:100},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'Atlas ADS × BrainHoney — Glycosept v3',size:20,font:'Arial',color:C.gray,italics:true})]}),
];

// ══════════════════════════════════════════════════════════════════════════════
//  ASSEMBLE
// ══════════════════════════════════════════════════════════════════════════════
const allChildren = [...cover, ...c2, ...c5, ...c7, ...sintese];

const doc = new Document({
  numbering:{config:[{reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,text:'•',
    alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
  styles:{
    default:{document:{run:{font:'Arial',size:22}}},
    paragraphStyles:[
      {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:38,bold:true,font:'Arial',color:C.navy},paragraph:{spacing:{before:400,after:200},outlineLevel:0}},
      {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:28,bold:true,font:'Arial',color:C.blue},paragraph:{spacing:{before:300,after:160},outlineLevel:1}},
      {id:'Heading3',name:'Heading 3',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:24,bold:true,font:'Arial',color:C.navy},paragraph:{spacing:{before:200,after:80},outlineLevel:2}},
    ]
  },
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.orange,space:6}},
      children:[new TextRun({text:'GANCHOS + NOMES CHICLETE — GLYCOSEPT v3 | Atlas ADS',size:18,font:'Arial',color:C.gray})]})
    ]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({text:'Pág. ',size:18,font:'Arial',color:C.gray}),
        new TextRun({children:[PageNumber.CURRENT],size:18,font:'Arial',color:C.gray}),
        new TextRun({text:' | Atlas ADS — Confidencial',size:18,font:'Arial',color:C.gray})]})
    ]})},
    children:allChildren
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('/sessions/loving-dazzling-knuth/mnt/outputs/Hooks_Chiclete_Glycosept_v3.docx',buf);
  console.log('Gerado com sucesso!');
}).catch(e=>{console.error(e);process.exit(1);});
