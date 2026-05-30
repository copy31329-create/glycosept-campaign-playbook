const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, Header, Footer, PageBreak
} = require('docx');
const fs = require('fs');

// ─── PALETTE ────────────────────────────────────────────────────────────────
const C = {
  navy:   '1F3864', blue:  '2E75B6', orange: 'C55A11',
  red:    'C00000', green: '375623', gray:   '595959',
  lgray:  'F2F2F2', white: 'FFFFFF', gold:   'BF8F00',
  teal:   '1D6C6C', purple:'5B2C8D'
};

// Concept accent colours (one per concept)
const ACCENTS = [C.orange, C.teal, C.red, C.purple, C.blue, C.green, C.gold];

const CW = 9360;
const CELL_M = { top:100, bottom:100, left:140, right:140 };
const BL = { style: BorderStyle.SINGLE, size:1, color:'CCCCCC' };
const BORDERS = { top:BL, bottom:BL, left:BL, right:BL };

// ─── HELPERS ────────────────────────────────────────────────────────────────
function h1(t, color=C.navy) {
  return new Paragraph({ heading:HeadingLevel.HEADING_1, spacing:{before:400,after:200},
    children:[new TextRun({text:t,bold:true,size:38,color,font:'Arial'})] });
}
function h2(t, color=C.blue) {
  return new Paragraph({ heading:HeadingLevel.HEADING_2, spacing:{before:300,after:160},
    children:[new TextRun({text:t,bold:true,size:28,color,font:'Arial'})] });
}
function h3(t, color=C.navy) {
  return new Paragraph({ heading:HeadingLevel.HEADING_3, spacing:{before:220,after:100},
    children:[new TextRun({text:t,bold:true,size:24,color,font:'Arial'})] });
}
function label(t, color=C.orange) {
  return new Paragraph({ spacing:{before:160,after:60},
    children:[new TextRun({text:`▸ ${t}`,bold:true,size:22,color,font:'Arial'})] });
}
function p(text, extra={}) {
  return new Paragraph({ spacing:{before:60,after:80},
    children:[new TextRun({text,size:22,font:'Arial',...extra})] });
}
function pB(text, color=C.navy) {
  return p(text,{bold:true,color});
}
function pR(runs) {
  return new Paragraph({ spacing:{before:60,after:80},
    children:runs.map(r=>new TextRun({size:22,font:'Arial',...r})) });
}
function hook(text, accent=C.orange) {
  return new Paragraph({
    spacing:{before:120,after:120}, indent:{left:720},
    border:{left:{style:BorderStyle.SINGLE,size:18,color:accent,space:10}},
    shading:{fill:'FFF2E8',type:ShadingType.CLEAR},
    children:[new TextRun({text,size:24,font:'Arial',bold:true,color:C.navy})]
  });
}
function hookVar(text, accent=C.blue) {
  return new Paragraph({
    spacing:{before:80,after:80}, indent:{left:720},
    border:{left:{style:BorderStyle.SINGLE,size:8,color:accent,space:10}},
    children:[new TextRun({text:`"${text}"`,size:22,font:'Arial',italics:true,color:C.gray})]
  });
}
function pill(text, fill=C.orange) {
  return new Paragraph({
    spacing:{before:200,after:100},
    children:[new TextRun({text:`  ${text}  `,bold:true,size:20,font:'Arial',color:C.white,
      shading:{fill,type:ShadingType.CLEAR}})]
  });
}
function divider(color=C.blue) {
  return new Paragraph({ spacing:{before:240,after:240},
    border:{bottom:{style:BorderStyle.SINGLE,size:8,color,space:1}}, children:[] });
}
function bullet(text, bold_prefix='') {
  const ch = bold_prefix
    ? [new TextRun({text:bold_prefix+' ',bold:true,size:22,font:'Arial'}),
       new TextRun({text,size:22,font:'Arial'})]
    : [new TextRun({text,size:22,font:'Arial'})];
  return new Paragraph({ spacing:{before:40,after:40},
    numbering:{reference:'bullets',level:0}, children:ch });
}
function pb() { return new Paragraph({children:[new PageBreak()]}); }

function tbl(rows, colWidths, hFill=C.blue) {
  return new Table({
    width:{size:CW,type:WidthType.DXA}, columnWidths:colWidths,
    rows:rows.map((row,ri)=>new TableRow({
      tableHeader:ri===0,
      children:row.map((cell,ci)=>new TableCell({
        borders:BORDERS,
        width:{size:colWidths[ci],type:WidthType.DXA},
        margins:CELL_M,
        shading:ri===0?{fill:hFill,type:ShadingType.CLEAR}:{fill:C.white,type:ShadingType.CLEAR},
        verticalAlign:VerticalAlign.TOP,
        children:Array.isArray(cell)?cell:[new Paragraph({spacing:{before:40,after:40},
          children:[new TextRun({text:cell,size:20,font:'Arial',bold:ri===0,color:ri===0?C.white:C.navy})]})]
      }))
    }))
  });
}

// Concept header block
function conceptHeader(num, name, accent) {
  return [
    pb(),
    new Paragraph({
      spacing:{before:0,after:0},
      shading:{fill:accent,type:ShadingType.CLEAR},
      children:[new TextRun({text:`  CONCEITO ${num} — ${name}  `,bold:true,size:40,color:C.white,font:'Arial'})]
    }),
    new Paragraph({ spacing:{before:0,after:300}, shading:{fill:accent,type:ShadingType.CLEAR}, children:[] }),
  ];
}

// Structured row for the quick-ref table
function qrow(field, value) {
  return [field, value];
}

// ══════════════════════════════════════════════════════════════════════════════
//  COVER
// ══════════════════════════════════════════════════════════════════════════════
const cover = [
  new Paragraph({ spacing:{before:2400,after:300}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'7 CONCEITOS DE ADS DE ALTA PERFORMANCE',bold:true,size:44,color:C.navy,font:'Arial'})] }),
  new Paragraph({ spacing:{before:0,after:200}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'GLYCOSEPT — DIABETES TIPO 2',bold:true,size:56,color:C.orange,font:'Arial'})] }),
  new Paragraph({ spacing:{before:100,after:100}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'Mecanismos Novos · Nomes Chiclete · Hooks por Plataforma',size:24,font:'Arial',color:C.gray,italics:true})] }),
  divider(C.orange),
  new Paragraph({ spacing:{before:300,after:100}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'Facebook Ads  ·  TikTok Ads  ·  Telegram',bold:true,size:26,font:'Arial',color:C.navy})] }),
  new Paragraph({ spacing:{before:200,after:600}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:`Gerado em ${new Date().toLocaleDateString('pt-BR')} | Atlas ADS — Confidencial`,size:20,font:'Arial',color:C.gray})] }),
  tbl([
    ['#','Conceito','Mecanismo','Emoção-Chave','Plataforma Ideal'],
    ['1','O Código do Urso',      'Código Metabólico Selvagem',  'Espanto + Esperança',     'FB · TikTok'],
    ['2','O Interruptor Apagado', 'Blackout Metabólico',         'Revelação + Indignação',  'FB · TikTok'],
    ['3','O Hormônio Roubado',    'Sequestro Hormonal',          'Raiva + Alívio',          'FB · Telegram'],
    ['4','Lama Doce',             'Entupimento Visceral',        'Nojo + Urgência',         'TikTok · FB'],
    ['5','O Espelho de 2 Anos',   'Identidade Sequestrada',      'Tristeza → Esperança',    'FB · Telegram'],
    ['6','A Sentinela Cega',      'Vigilância Metabólica',       'Curiosidade + Controle',  'TikTok · FB'],
    ['7','O Gêmeo Saudável',      'DNA Dividido',                'Injustiça → Possibilidade','FB · Native'],
  ],[400,2000,2300,2200,2460], C.navy),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 1 — O CÓDIGO DO URSO
// ══════════════════════════════════════════════════════════════════════════════
const c1 = [
  ...conceptHeader(1,'O CÓDIGO DO URSO', ACCENTS[0]),

  h2('① NOME DO MECANISMO', C.orange),
  pB('"Código Metabólico Selvagem"'),
  p('O mecanismo posiciona o GLP-1 não como um hormônio técnico, mas como um "código genético ancestral" que ursos evoluíram durante 10.000 anos para sobreviver à hibernação — e que existe dormido no DNA humano, bloqueado pela lama tóxica ambiental. O produto "reativa esse código".'),

  h2('② EXPLICAÇÃO PSICOLÓGICA DO MECANISMO', C.orange),
  p('A psicologia aqui opera em três camadas simultâneas:'),
  bullet('ESPANTO BIOLÓGICO: o urso come 20.000 calorias de mel puro e acorda sem diabetes — fato verificável que quebra toda crença prévia sobre açúcar e metabolismo.'),
  bullet('PERTENCIMENTO NATURAL: "esse código existe no seu DNA também" — cria inclusão. Não é algo externo. É algo que você perdeu e pode recuperar.'),
  bullet('CULPA TRANSFERIDA: "foi bloqueado por toxinas ambientais" — libera totalmente a responsabilidade do paciente. Não é fraqueza. É exposição.'),
  p('O "código" como metáfora é tecnológico o suficiente para soar científico, mas natural o suficiente para não intimidar o público 50+. Funciona como ponte entre natureza e ciência.'),

  h2('③ NOME CHICLETE DA OFERTA', C.orange),
  pR([{text:'Nome principal: ',bold:true},{text:'"Protocolo Grizzly"'}]),
  pR([{text:'Variações: ',bold:true},{text:'"Código Urso" · "Ativador Selvagem" · "Fórmula Grizzly" · "DNA do Açúcar Perfeito"'}]),
  p('"Protocolo Grizzly" cola porque combina autoridade científica (protocolo) com imagem de força e saúde selvagem (grizzly). É visual, é masculino e feminino ao mesmo tempo, e ninguém no mercado de diabetes usou esse nome.'),

  h2('④ HOOK PRINCIPAL', C.orange),
  hook('Um urso grizzly come 20.000 calorias de mel puro toda semana antes de hibernar. Seis meses sem comer. E acorda com o açúcar perfeito — zero diabetes, zero resistência à insulina. Os cientistas descobriram o código genético por trás disso. E ele existe no seu DNA também — só está bloqueado.', ACCENTS[0]),

  h2('⑤ VARIAÇÕES DE HOOKS', C.orange),
  p('Para Facebook (feed de notícia — curiosidade científica):'),
  hookVar('Por que ursos nunca ficam diabéticos? Comem açúcar puro o dia inteiro e acordam da hibernação com o metabolismo de um atleta. A ciência finalmente explicou. E a resposta muda tudo que você sabe sobre diabetes.', C.orange),
  hookVar('Eles tentaram criar um remédio a partir do hormônio dos ursos. Chamaram de Ozempic. Cobram R$2.000 por mês. O original — que os ursos usam — ainda funciona 3x melhor e não tem efeito colateral.', C.orange),
  p(''),
  p('Para TikTok (texto na tela, 0–3 segundos):'),
  hookVar('POV: você descobre que ursos têm a cura do diabetes e ninguém te contou', C.teal),
  hookVar('Por que ursos comem mel puro e NUNCA ficam diabéticos? 🐻', C.teal),
  hookVar('Estudo chocou a comunidade médica: urso grizzly = solução do diabetes?', C.teal),
  p(''),
  p('Para Telegram (formato "descoberta vazada"):'),
  hookVar('🐻 Isso foi suprimido por 3 anos: o hormônio que ursos produzem 5.000x mais que humanos é o que reverte o diabetes — e existe em 4 alimentos naturais...', C.blue),

  h2('⑥ IDEIA VISUAL DO CRIATIVO', C.orange),
  tbl([
    ['Plataforma','Formato','Descrição Visual'],
    ['Facebook / YouTube','VSL 60–90s',
      'Abre com B-roll de urso devorando colmeia de mel (footage de documentário, gratuito). Texto sobreposto: "20.000 calorias de mel puro". Corte para pesquisador em laboratório: "GLP-1: 5.000x". Corte para gota de mel em câmera lenta dourada. Texto: "4 ingredientes. 12 semanas. 97% reverteram." CTA aparesce sobre mel escorrendo.'],
    ['TikTok','15–30s vertical',
      'Abre com footage de urso + texto piscando "POR QUE ELE NUNCA FICA DIABÉTICO?". Áudio trending (curiosidade). Cuts rápidos: mel → laboratório → gráfico A1c descendo → depoimento 5 segundos → CTA. Parece um TikTok educativo de natureza, não anúncio.'],
    ['Telegram','Imagem estática + texto longo',
      'Foto de urso em floresta + mel + fundo escuro. Texto: "O hormônio que os ursos produzem 5.000x mais que você — e como replicar em 21 dias". Link abaixo.'],
  ],[1800,1800,5760], C.navy),

  h2('⑦ EMOÇÃO PRINCIPAL EXPLORADA', C.orange),
  pR([{text:'Espanto → Curiosidade → Esperança',bold:true,color:C.orange}]),
  p('A sequência emocional: (1) espanto com o paradoxo biológico do urso, (2) curiosidade sobre "como replicar", (3) esperança de que existe uma solução natural e poderosa que a ciência já validou. Não começa com medo ou dor — começa com maravilhamento. Isso é raro no nicho de diabetes e cria menos resistência inicial.'),

  h2('⑧ CTA IDEAL', C.orange),
  tbl([
    ['Plataforma','CTA','Tom'],
    ['Facebook','Assistir o Protocolo Grizzly Completo →','Curiosidade + exclusividade'],
    ['TikTok','Link na bio: Protocolo Grizzly 🐻','Orgânico, sem pressão'],
    ['Telegram','Ver antes que tirem do ar ↓','Urgência + supressão'],
    ['Email','Quero ativar meu Código Urso →','Identidade + pertencimento'],
  ],[2000,4000,3360], C.orange),

  h2('⑨ HEADLINE DA VSL', C.orange),
  hook('"O Hormônio que os Ursos Produzem 5.000 Vezes Mais que Você É a Única Coisa que Pode Reverter Seu Diabetes — e Existe em 4 Alimentos que Você Nunca Imaginou"', ACCENTS[0]),

  h2('⑩ ÂNGULO DE ESCALA', C.orange),
  p('O conceito "animal behavior as proof of mechanism" tem viralidade natural — pode ser compartilhado como conteúdo educativo sobre natureza e metabolismo sem parecer propaganda de produto. Funciona em cold traffic completamente frio (pessoas que nunca ouviram falar de Glycosept). É um conceito de topo de funil perfeito que alimenta retargeting com a VSL completa.'),
  p('Ângulo de escala: criar 5 variações do creative com espécies diferentes (urso → abelha melífera → tartaruga → tubarão-baleia → elefante) cada uma ligada a um dos 4 ingredientes. Campanha "A Natureza Já Resolveu".'),

  h2('⑪ COMO TORNAR MAIS NATIVO', C.orange),
  bullet('Usar footage real de documentário (BBC Earth, National Geographic — B-roll disponível) sem qualquer produção adicional visível.'),
  bullet('Sem Dr. com jaleco nos primeiros 30 segundos — apenas narrador over footage de natureza.'),
  bullet('Começar o áudio com som ambiente de floresta + abelhas — pattern interrupt no feed de saúde.'),
  bullet('Caption no Facebook: não começar com "Você sabia que..." — começar com a cena: "Um urso em Yellowstone acabou de comer 20.000 calorias de mel. Amanhã ele vai hibernar por 6 meses. E acordar sem diabetes. Aqui está o que os pesquisadores descobriram."'),
  bullet('TikTok: usar template de "fato de natureza chocante" — formato que já tem alto orgânico.'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 2 — O INTERRUPTOR APAGADO
// ══════════════════════════════════════════════════════════════════════════════
const c2 = [
  ...conceptHeader(2,'O INTERRUPTOR APAGADO', ACCENTS[1]),

  h2('① NOME DO MECANISMO', C.teal),
  pB('"Blackout Metabólico"'),
  p('O mecanismo reposiciona o diabetes não como falha do pâncreas, mas como um "blackout" — o sinal elétrico (GLP-1) que comanda as fábricas de insulina foi cortado pela lama tóxica. O pâncreas está intacto. O problema é o sinal. E sinais podem ser religados.'),

  h2('② EXPLICAÇÃO PSICOLÓGICA DO MECANISMO', C.teal),
  bullet('ESPERANÇA ESTRUTURAL: "o pâncreas está intacto" é a mensagem mais esperançosa possível para um diabético que foi told "seu pâncreas está destruído". Cria abertura emocional imediata.'),
  bullet('ANALOGIA COTIDIANA: o disjuntor elétrico é universalmente compreensível — não exige vocabulário médico. Todo mundo já teve luz cortada por um disjuntor desarmado. É uma experiência física familiar.'),
  bullet('AGÊNCIA RECUPERADA: "sinais podem ser religados" coloca o leitor de volta no controle. Não é uma sentença — é um problema técnico com solução técnica.'),
  bullet('INVERSÃO DA NARRATIVA MÉDICA: "seu médico está tratando o sintoma (açúcar) não o problema (sinal cortado)" cria indignação controlada que motiva ação.'),

  h2('③ NOME CHICLETE DA OFERTA', C.teal),
  pR([{text:'Nome principal: ',bold:true},{text:'"Protocolo Liga-Desliga" ou "Interruptor Pancreático"'}]),
  pR([{text:'Variações: ',bold:true},{text:'"Relê Metabólico" · "Código de Religação" · "Switch do Açúcar" · "Fórmula Blackout"'}]),
  p('"Liga-Desliga" é simples, visual e imediatamente compreensível. "Blackout Metabólico" funciona como nome do mecanismo do problema — carregado, urgente, impossível de esquecer.'),

  h2('④ HOOK PRINCIPAL', C.teal),
  hook('Seu pâncreas não parou de funcionar. Alguém cortou a luz. O problema não é o órgão — é o sinal que foi bloqueado por dentro. E é por isso que metformin, dieta e até Ozempic nunca resolvem: eles tratam o sintoma. Ninguém religou o interruptor.', ACCENTS[1]),

  h2('⑤ VARIAÇÕES DE HOOKS', C.teal),
  p('Facebook (indignação + revelação):'),
  hookVar('Seu médico está tratando o açúcar no sangue. O problema real é que o sinal que comanda seu pâncreas foi cortado. É a diferença entre trocar uma lâmpada queimada e religar o disjuntor. Um funciona. O outro, não.', C.teal),
  hookVar('Imagina que sua casa está escura. Não porque a energia acabou — porque um disjuntor desarmou. É exatamente isso que acontece no seu pâncreas com o diabetes tipo 2. E o interruptor pode ser religado.', C.teal),
  p(''),
  p('TikTok (visual + chocante, texto na tela):'),
  hookVar('Seu pâncreas não está quebrado. O sinal foi cortado 🔌', C.teal),
  hookVar('Por que metformin não cura? Porque trata o sintoma, não o interruptor ⚡', C.teal),
  hookVar('Diabetes tipo 2 = blackout metabólico. Isso muda TUDO sobre tratamento.', C.teal),
  p(''),
  p('Telegram (formato urgente/informativo):'),
  hookVar('⚡ Isso explica por que seu açúcar não estabiliza: não é o pâncreas que falhou — é o sinal que foi bloqueado. E existe um protocolo de 4 ingredientes que restabelece esse sinal em 21 dias...', C.blue),

  h2('⑥ IDEIA VISUAL DO CRIATIVO', C.teal),
  tbl([
    ['Plataforma','Formato','Descrição Visual'],
    ['Facebook','Imagem estática / carrossel',
      'Frame 1: quadro elétrico com disjuntor desarmado, luz vermelha. Texto: "É assim que está seu pâncreas agora." Frame 2: mão religando o disjuntor. Texto: "É isso que 4 ingredientes naturais fazem." Frame 3: A1c numbers caindo. CTA.'],
    ['TikTok','15s animação simples',
      'Animação: pâncreas cinza com sinal vermelho (interrompido) → protocolo é aplicado → sinal verde → insulina fluindo → glicemia estabilizada. Simples, low-fi, parece content educativo. Áudio: som de eletricidade + "clique" do disjuntor religando.'],
    ['YouTube Pre-roll','30s talking head',
      'Pessoa comum (não Dr.) na frente de um quadro elétrico real: "Você sabia que diabetologistas estão tratando a lâmpada quando o problema é o disjuntor?" [liga o disjuntor, luzes acendem] "Aqui está o que religar no seu corpo."'],
  ],[1800,1800,5760], C.teal),

  h2('⑦ EMOÇÃO PRINCIPAL', C.teal),
  pR([{text:'Revelação → Indignação → Esperança Técnica',bold:true,color:C.teal}]),
  p('A sequência emocional começa com revelação ("meu médico está tratando a coisa errada"), evolui para indignação controlada ("11 anos de tratamento errado") e chega à esperança técnica ("se é um sinal, pode ser religado"). O público sente que finalmente entendeu o problema de verdade — e isso cria alta intenção de ação.'),

  h2('⑧ CTA IDEAL', C.teal),
  tbl([
    ['Plataforma','CTA'],
    ['Facebook / YouTube','Ver Como Religar o Interruptor →'],
    ['TikTok','Clique no link da bio para ver o protocolo completo ⚡'],
    ['Telegram','Protocolo de Religação: [link] (funciona em 21 dias)'],
  ],[2500,6860], C.teal),

  h2('⑨ HEADLINE DA VSL', C.teal),
  hook('"Seu Médico Está Trocando a Lâmpada Quando o Problema É o Disjuntor — e Aqui Está o Protocolo que Finalmente Religar o Sinal que Comanda Seu Pâncreas"', ACCENTS[1]),

  h2('⑩ ÂNGULO DE ESCALA', C.teal),
  p('O ângulo elétrico escala horizontalmente para múltiplos problemas de saúde (pressão, tireoide, fígado) com o mesmo mecanismo-base de "sinal bloqueado". Pode virar uma linha de produtos com o mesmo frame conceitual. Para Glycosept especificamente: criar série de 3 vídeos curtos — "O Problema", "O Interruptor", "O Protocolo" — sequência perfeita para retargeting escalonado.'),

  h2('⑪ COMO TORNAR MAIS NATIVO', C.teal),
  bullet('Usar footage de eletricista real consertando quadro elétrico — muito mais nativo do que animação médica.'),
  bullet('Caption começa com experiência cotidiana: "Minha casa ficou sem luz na semana passada por causa de um disjuntor desarmado. Enquanto chamava o eletricista, percebi que era exatamente isso que acontecia com meu pâncreas há 8 anos."'),
  bullet('TikTok: duet format — reagir a um vídeo de eletricista consertando disjuntor, aplicando a metáfora ao diabetes. Pattern interrupt de nicho cruzado.'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 3 — O HORMÔNIO ROUBADO
// ══════════════════════════════════════════════════════════════════════════════
const c3 = [
  ...conceptHeader(3,'O HORMÔNIO ROUBADO', ACCENTS[2]),

  h2('① NOME DO MECANISMO', C.red),
  pB('"Sequestro Hormonal"'),
  p('As farmacêuticas descobriram que o GLP-1 — hormônio que o corpo produz naturalmente — reverte o diabetes. Em vez de ensinar as pessoas a produzi-lo (gratuito, sem efeito colateral), sintetizaram uma versão que cria dependência eterna. Ozempic é a cópia sintética do seu próprio hormônio, com prazo de validade e fatura mensal. O "original" existe em 4 ingredientes naturais.'),

  h2('② EXPLICAÇÃO PSICOLÓGICA DO MECANISMO', C.red),
  bullet('RAIVA MORAL JUSTA: "te roubaram um hormônio que é seu" — não é ceticismo geral contra farmácias. É raiva específica e pessoal. Altamente motivadora para ação imediata.'),
  bullet('DADO IRREFUTÁVEL COMO ÂNCORA: 67% das pessoas que param o Ozempic recuperam tudo — verificável, real, e humilhante para as farmacêuticas. Destrói o argumento "mas Ozempic funciona."'),
  bullet('EXCLUSIVIDADE INVERTIDA: em vez de "conteúdo exclusivo", é "eles não queriam que você soubesse." A exclusividade do segredo aumenta o valor percebido da informação.'),
  bullet('URGÊNCIA MORAL: compartilhar esse conteúdo parece um ato de resistência contra um sistema corrupto — o que aumenta organicamente o share rate e a viralidade.'),

  h2('③ NOME CHICLETE DA OFERTA', C.red),
  pR([{text:'Nome principal: ',bold:true},{text:'"GLP-1 Original" ou "O Hormônio Roubado" ou "Ozempic Natural"'}]),
  pR([{text:'Variações: ',bold:true},{text:'"Protocolo Anti-Sequestro" · "Hormônio Livre" · "Código GLP" · "A Versão Original"'}]),
  p('"Ozempic Natural" é o nome mais poderoso porque ancora diretamente na referência cultural de mercado — mas sem os $2.000/mês e sem agulha. É uma comparação embutida no nome.'),

  h2('④ HOOK PRINCIPAL', C.red),
  hook('A Pfizer sabe que um hormônio que você produz naturalmente reverte o diabetes. Em vez de te ensinar a produzi-lo — de graça, sem efeito colateral — eles criaram uma versão sintética que custa R$2.000 por mês e te deixa dependente para sempre. 67% de quem para o Ozempic recupera tudo. Aqui está o hormônio original.', ACCENTS[2]),

  h2('⑤ VARIAÇÕES DE HOOKS', C.red),
  p('Facebook (conspiração + dado verificável):'),
  hookVar('Ozempic não cura diabetes. É uma cópia sintética do seu próprio hormônio — que te força a depender de uma injeção de R$2.000 por mês para o resto da vida. O original existe em 4 alimentos. Sem agulha. Sem prazo de validade. Sem dependência.', C.red),
  hookVar('67% de quem para o Ozempic recupera todo o peso perdido. Porque é uma cópia que cria dependência. O hormônio original — que seu corpo deveria produzir naturalmente — ainda existe. E aqui está onde encontrá-lo.', C.red),
  p(''),
  p('TikTok (choque + simplicidade):'),
  hookVar('O Ozempic é uma CÓPIA do seu próprio hormônio 😱 (e o original é gratuito)', C.red),
  hookVar('Por que 67% recuperam o peso ao parar o Ozempic? Porque é uma cópia, não o original 🧬', C.red),
  hookVar('Big Pharma sequestrou seu hormônio e te cobra R$2.000/mês pela cópia sintética', C.red),
  p(''),
  p('Telegram (formato "vazamento"):'),
  hookVar('🚨 Isso explica por que o Ozempic funciona enquanto você toma e para quando você para: ele injeta uma cópia sintética do GLP-1 — o hormônio que seu corpo deveria produzir naturalmente. A versão original existe em 4 compostos naturais. Sem agulha. Sem conta mensal. O vídeo completo está aqui antes que removam:', C.blue),

  h2('⑥ IDEIA VISUAL DO CRIATIVO', C.red),
  tbl([
    ['Plataforma','Formato','Descrição Visual'],
    ['Facebook','Comparativo visual',
      '"CÓPIA vs. ORIGINAL" — dois lados: esquerda: seringa Ozempic + R$2.000/mês + lista de efeitos colaterais (vermelho). Direita: pote de mel Manuka + abelha + "R$X/mês" + "zero efeitos colaterais" (verde). Tabela simples, alta legibilidade, funciona como imagem estática ou carrossel.'],
    ['TikTok','15-30s talking head',
      'Pessoa segura caixa de Ozempic em uma mão e pote de mel na outra. "Um desses custa R$2.000 por mês. O outro custa X. Ambos ativam o mesmo hormônio. Mas só um te deixa dependente para sempre." Coloca o Ozempic no chão. Abre o mel.'],
    ['Telegram','Screenshot + texto',
      'Screenshot editado de artigo científico (real) sobre GLP-1 + texto sobreposto: "Eles sabiam. Escolheram te cobrar R$2.000/mês pela cópia." Link abaixo.'],
  ],[1800,1800,5760], C.red),

  h2('⑦ EMOÇÃO PRINCIPAL', C.red),
  pR([{text:'Raiva → Indignação Moral → Alívio (existe saída)',bold:true,color:C.red}]),
  p('A raiva funciona como combustível de ação imediata — o público diabético que gasta fortunas em medicação está pré-carregado com frustração. Este conceito valida essa frustração, direciona para um inimigo externo concreto (Big Pharma) e imediatamente oferece a alternativa. É a sequência emocional de maior conversão em DR: raiva → inimigo → solução.'),

  h2('⑧ CTA IDEAL', C.red),
  tbl([
    ['Plataforma','CTA'],
    ['Facebook','Ver o Hormônio Original →'],
    ['TikTok','Protocolo completo no link da bio 🧬'],
    ['Telegram','O original está aqui: [link]'],
  ],[2500,6860], C.red),

  h2('⑨ HEADLINE DA VSL', C.red),
  hook('"A Big Pharma Sintetizou Seu Hormônio Anti-Diabético Natural, Cobrou R$2.000 por Mês Pela Cópia e Nunca Te Contou que o Original Ainda Existe — Aqui Está Onde Encontrá-lo"', ACCENTS[2]),

  h2('⑩ ÂNGULO DE ESCALA', C.red),
  p('Este é o conceito com maior potencial viral orgânico — indignação contra Big Pharma gera compartilhamento espontâneo em grupos de Facebook, WhatsApp e Telegram sem investimento adicional. O ângulo escala para qualquer produto "natural vs. sintético" no mercado de saúde. Para Glycosept: criar variações para cada medicamento (metformin, glipizide, insulina) com o mesmo frame "cópia vs. original".'),

  h2('⑪ COMO TORNAR MAIS NATIVO', C.red),
  bullet('Facebook: formato de "post pessoal indignado" — sem logo de empresa, sem jaleco, sem CTA explícito nos primeiros 3 parágrafos. Parece alguém desabafando sobre o sistema de saúde.'),
  bullet('TikTok: usar trending sound de "exposé" ou jornalismo investigativo. Caption: "Coisas que a indústria farmacêutica não quer que você saiba #diabetes #ozempic #saude"'),
  bullet('Telegram: formato de "mensagem encaminhada" — parece que alguém descobriu algo e está compartilhando com amigos, não um anúncio pago.'),
  bullet('Adicionar call-out específico: "Se você toma Ozempic agora, não pare antes de ver isso — tem uma forma de sair da dependência com segurança."'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 4 — LAMA DOCE
// ══════════════════════════════════════════════════════════════════════════════
const c4 = [
  ...conceptHeader(4,'LAMA DOCE', ACCENTS[3]),

  h2('① NOME DO MECANISMO', C.purple),
  pB('"Entupimento Pancreático" / "Lama Doce"'),
  p('O paradoxo central: o próprio açúcar que você tenta evitar cria, quando combinado com toxinas ambientais, uma lama viscosa que entope o pâncreas — impedindo que ele processe qualquer açúcar, independente da dieta. O nome "Lama Doce" é o paradoxo perfeito: o que parece inofensivo (doçura) é o que sufoca.'),

  h2('② EXPLICAÇÃO PSICOLÓGICA DO MECANISMO', C.purple),
  bullet('NOJO COMO MOTIVADOR: a imagem de lama viscosa por dentro é visceral e repulsiva — ativa urgência de "limpeza" que vai além do racional. O público quer dissolver a lama, não apenas controlar o açúcar.'),
  bullet('METÁFORA DOMÉSTICA UNIVERSAL: cano entupido com gordura é a experiência mais familiar possível. Todo mundo já viu, todo mundo já sentiu o cano lento. O pâncreas entupido é imediatamente compreensível sem vocabulário médico.'),
  bullet('PARADOXO DE CULPA TRANSFERIDA: "não é o açúcar que você come, é o que o açúcar deixa para trás com as toxinas" — remove a culpa da dieta e a transfere para o ambiente. Libera a resignação.'),
  bullet('DEMO VIRÁVEL: a demonstração da dissolução (substância viscosa num funil sendo dissolvida pelo produto) é o conteúdo mais virável do nicho de saúde — ciência visual que parece experimento, não propaganda.'),

  h2('③ NOME CHICLETE DA OFERTA', C.purple),
  pR([{text:'Nome principal: ',bold:true},{text:'"Dissolvente Pancreático" ou "Fórmula Limpeza Total" ou "Protocolo Dissolução"'}]),
  pR([{text:'Variações: ',bold:true},{text:'"Anti-Lama" · "Desentupidor Interno" · "Flush Metabólico" · "Limpeza do Pâncreas"'}]),
  p('"Desentupidor Interno" é provocativo e memorável — soa quase rude, o que o torna impossível de esquecer. "Flush Metabólico" tem apelo mais clean/premium. Escolha depende do tom da campanha.'),

  h2('④ HOOK PRINCIPAL', C.purple),
  hook('Você controla o açúcar que come. Mas existe uma lama tóxica que se acumula no seu pâncreas independente da sua dieta — e ela está lá agora, sufocando as células que produzem insulina. Não é o que você come. É o que ficou preso por dentro.', ACCENTS[3]),

  h2('⑤ VARIAÇÕES DE HOOKS', C.purple),
  p('Facebook (visual + urgência):'),
  hookVar('Do mesmo jeito que gordura entope um cano, existe uma substância tóxica entopindo seu pâncreas agora. E ela não sai com dieta, metformin ou Ozempic. Porque nenhum deles foi feito para dissolvê-la.', C.purple),
  hookVar('Seu pâncreas tem um entupimento. É uma lama formada por toxinas do ambiente — pesticidas, microplásticos, BPA — que se acumula por anos sem sintoma visível. Até o dia em que você recebe o diagnóstico. Aqui está o que dissolve.', C.purple),
  p(''),
  p('TikTok (demo + choque):'),
  hookVar('Isso é o que está no seu pâncreas agora 🤢 (e como dissolver)', C.purple),
  hookVar('Por que dieta NÃO cura diabetes? Porque o problema é essa lama aqui 👇', C.purple),
  hookVar('Experimento ao vivo: dissolvendo a "lama pancreática" com 4 ingredientes naturais', C.purple),
  p(''),
  p('Telegram:'),
  hookVar('🧫 Pesquisadores de Harvard nomearam isso de "lama pancreática" — um depósito tóxico que se forma no pâncreas de 90% dos adultos acima de 50 anos. Aqui está o protocolo que dissolve em 21 dias:', C.blue),

  h2('⑥ IDEIA VISUAL DO CRIATIVO', C.purple),
  tbl([
    ['Plataforma','Formato','Descrição Visual'],
    ['Facebook / YouTube','Demo 30-60s',
      'Câmera de laboratório. Funil de vidro transparente. Pesquisador despeja substância escura e viscosa (representa sludge). Líquido (insulina) tenta passar — passa lento, conta-gotas. Adiciona mistura dourada (mel Manuka). Substância viscosa dissolve. Líquido flui livre. Texto: "É isso que acontece no seu pâncreas quando você toma Glycosept." Altíssimo potential de parar o scroll.'],
    ['TikTok','15s experimento',
      'Hands-only, sem rosto. Copo com geleia preta/escura. Adiciona líquido dourado. Dissolução em câmera lenta. Texto na tela: "Seu pâncreas → 4 ingredientes naturais → Seu pâncreas 21 dias depois". Áudio: satisfying sound de dissolução. Parece um TikTok de experimento viral, não anúncio.'],
    ['Imagem Estática','Comparativo visual',
      'À esquerda: cano de encanamento entupido com gordura escura. À direita: pâncreas com lama. Texto central: "Mesma coisa. Diferente lugar." CTA abaixo.'],
  ],[1800,1800,5760], C.purple),

  h2('⑦ EMOÇÃO PRINCIPAL', C.purple),
  pR([{text:'Nojo/Urgência → Curiosidade (como dissolver?) → Alívio (existe solução)',bold:true,color:C.purple}]),
  p('O nojo é subestimado em copy de saúde. Quando aplicado a algo que está "dentro do corpo", cria urgência de limpeza imediata — mais visceral e motivadora do que medo de complicações futuras. "Tem lama no meu pâncreas agora" é mais urgente do que "daqui a 5 anos posso ter problemas renais."'),

  h2('⑧ CTA IDEAL', C.purple),
  tbl([
    ['Plataforma','CTA'],
    ['Facebook','Ver o Protocolo de Dissolução Completo →'],
    ['TikTok','Link na bio: Dissolução em 21 dias 🧫'],
    ['Telegram','Protocolo Anti-Lama: [link]'],
  ],[2500,6860], C.purple),

  h2('⑨ HEADLINE DA VSL', C.purple),
  hook('"A Lama Tóxica que Está Sufocando Seu Pâncreas Independente da Sua Dieta — e o Protocolo de 4 Ingredientes que a Dissolve em 21 Dias"', ACCENTS[3]),

  h2('⑩ ÂNGULO DE ESCALA', C.purple),
  p('A demo visual escala para qualquer formato — short-form, long-form, imagem estática, GIF animado. Pode virar uma série de "experimentos" com outros ingredientes do produto. O conceito "dissolução" é transferível para outros nichos (fígado, artérias, articulações). Para Glycosept: criar 3 variações da demo com diferentes intensidades de "lama" para representar estágios do diabetes.'),

  h2('⑪ COMO TORNAR MAIS NATIVO', C.purple),
  bullet('TikTok: filmar como um "experimento de curiosidade" sem mencionar produto até o segundo 20. Caption: "Tentei replicar o que acontece no pâncreas de um diabético em casa..."'),
  bullet('Facebook: postar a demo como "vídeo científico compartilhável" sem CTA nos primeiros 20 segundos. Parece conteúdo educativo de um canal de saúde.'),
  bullet('Usar palavras do senso comum, não médico: "entupimento" em vez de "acúmulo de toxinas lipídicas", "desentupir" em vez de "dissolução de depósitos tóxicos".'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 5 — O ESPELHO DE 2 ANOS
// ══════════════════════════════════════════════════════════════════════════════
const c5 = [
  ...conceptHeader(5,'O ESPELHO DE 2 ANOS', ACCENTS[4]),

  h2('① NOME DO MECANISMO', C.blue),
  pB('"Identidade Sequestrada"'),
  p('O diabetes não rouba apenas a saúde. Rouba quem você era. A "Identidade Sequestrada" nomeia o dano psicológico real que o VSL original acerta em cheio mas não nomeia explicitamente: o provedor que se tornou dependente, o avô ativo que se tornou paciente, o parceiro que virou fardo. O produto não oferece apenas controle glicêmico — oferece devolução da identidade.'),

  h2('② EXPLICAÇÃO PSICOLÓGICA DO MECANISMO', C.blue),
  bullet('DEEPER CORE DIRETO: enquanto outros anúncios falam de A1c e glicemia, este fala de quem a pessoa era e quer voltar a ser. É o único ângulo que toca o nível mais profundo da dor — identidade, não sintoma.'),
  bullet('BEFORE/AFTER DE IDENTIDADE (não de corpo): o contraste não é "gordo vs. magro" nem "A1c alto vs. baixo" — é "a pessoa que você era vs. o paciente que o diabetes fez de você". Mais poderoso porque é universal para qualquer diabético independente de peso ou aparência.'),
  bullet('ESPELHO COMO METÁFORA CENTRAL: o espelho é o lugar onde confrontamos nossa identidade. "Quando foi a última vez que você viu a pessoa que você costumava ser no espelho?" é a pergunta que ninguém quer fazer — e por isso não para de ressoar.'),
  bullet('PRESENÇA FAMILIAR COMO BENEFÍCIO MÁXIMO: o Deeper Core do público 50+ não é saúde — é presença. Estar presente para filhos, netos, cônjuge. O produto promete isso, não apenas glicemia controlada.'),

  h2('③ NOME CHICLETE DA OFERTA', C.blue),
  pR([{text:'Nome principal: ',bold:true},{text:'"Protocolo do Retorno" ou "Reversão de Identidade"'}]),
  pR([{text:'Variações: ',bold:true},{text:'"Volta pra Você" · "O Eu Original" · "Protocolo de Presença" · "Fórmula do Retorno"'}]),
  p('"Protocolo do Retorno" é emocional e aspiracional — soa como voltar para casa, voltar para si mesmo. "Reversão de Identidade" é mais clínico mas extremamente preciso para o que o produto realmente faz na percepção do público.'),

  h2('④ HOOK PRINCIPAL', C.blue),
  hook('Quando foi a última vez que você olhou no espelho e viu a pessoa que você costumava ser — não o diabético, não o paciente, não o fardo? Essa pessoa ainda está lá. O diabetes não a destruiu. Só a prendeu. E existe uma saída.', ACCENTS[4]),

  h2('⑤ VARIAÇÕES DE HOOKS', C.blue),
  p('Facebook (confissão emocional):'),
  hookVar('Minha filha me olhou com pena pela última vez há 8 semanas. Hoje ela me olha com orgulho. Não porque meu A1c mudou — porque eu voltei a ser quem eu era antes do diabetes.', C.blue),
  hookVar('Eu parei de me reconhecer. Não pela aparência. Por dentro. O diabetes não me roubou só a saúde — roubou quem eu era. Encontrei uma forma de me reconhecer de volta.', C.blue),
  hookVar('Tem uma diferença entre controlar o diabetes e se libertar dele. Eu só aprendi essa diferença depois de 11 anos de metformin que não resolvia. Aqui está o que mudou.', C.blue),
  p(''),
  p('TikTok (emocional + identificação):'),
  hookVar('POV: você tem diabetes e parou de se reconhecer no espelho 🪞', C.blue),
  hookVar('Ninguém fala sobre o que o diabetes faz com quem você ERA 💔', C.blue),
  hookVar('Meu neto me perguntou por que eu estava sempre cansado. Aqui está o que mudou tudo 👴→🏃', C.blue),
  p(''),
  p('Telegram (depoimento pessoal):'),
  hookVar('📱 Compartilhei isso com 3 pessoas que conheço com diabetes e todas usaram. A diferença não foi só nos exames — foi em quem elas voltaram a ser. Link do protocolo completo:', C.blue),

  h2('⑥ IDEIA VISUAL DO CRIATIVO', C.blue),
  tbl([
    ['Plataforma','Formato','Descrição Visual'],
    ['Facebook','Split screen emocional',
      'Esquerda (antes): pessoa exausta verificando glicemia às 3am, mão trêmula, olhos fundos. Direita (depois): mesma pessoa (mesmo ator/atriz) na mesa de jantar com família, rindo, sem aparelho na mesa. Texto: "Não é sobre A1c. É sobre quem você voltou a ser." SEM narração médica — apenas música emocional + depoimento de 10 segundos.'],
    ['TikTok','Talking head pessoal',
      'Pessoa comum, câmera de celular, luz natural. Começa: "Vou te contar a coisa mais difícil de admitir sobre ter diabetes..." [pausa] "Eu parei de me reconhecer." Sem cortes. Sem grafismo. Autêntico. Nos segundos 25-50: o protocolo em 3 frases. Nos segundos 50-60: CTA.'],
    ['Facebook','Carrossel de identidade',
      'Frame 1: "O diabetes roubou isso de você..." [lista: energia, espontaneidade, liberdade alimentar, olhar de orgulho da família]. Frame 2: "E pode devolver tudo isso." Frame 3: Depoimento de Margaret (A1c 9.4→5.2 + "ate birthday cake with my granddaughter"). Frame 4: CTA.'],
  ],[1800,1800,5760], C.blue),

  h2('⑦ EMOÇÃO PRINCIPAL', C.blue),
  pR([{text:'Tristeza Reconhecida → Identificação Profunda → Desejo de Restauração',bold:true,color:C.blue}]),
  p('Este é o único conceito que começa na tristeza — e funciona porque o público 50+ com diabetes crônico está em luto por uma identidade perdida. Nomear essa tristeza ("você parou de se reconhecer") cria identificação instantânea. Quem nunca pensou nisso começa a pensar. Quem já pensou sente que alguém finalmente entendeu. É o gatilho de conversão mais profundo de todos os 7 conceitos.'),

  h2('⑧ CTA IDEAL', C.blue),
  tbl([
    ['Plataforma','CTA'],
    ['Facebook','Quero me Reconhecer de Volta →'],
    ['TikTok','Protocolo do Retorno: link na bio 🪞'],
    ['Telegram','Aqui está o protocolo que devolve quem você era: [link]'],
  ],[2500,6860], C.blue),

  h2('⑨ HEADLINE DA VSL', C.blue),
  hook('"O Diabetes Não Roubou Só a Sua Saúde. Roubou Quem Você Era. Este Protocolo de 4 Ingredientes Devolve os Dois — em 90 Dias ou Você Não Paga Nada"', ACCENTS[4]),

  h2('⑩ ÂNGULO DE ESCALA', C.blue),
  p('Este conceito escala por segmentação de persona — criar variações para diferentes identidades perdidas: o pai/avô ativo (fishing story), a mãe/avó cuidadora (Margaret story), o cônjuge presente ("o marido/esposa que ela conheceu"). Cada variação fala para uma fatia diferente do mesmo público. Custo de produção baixo — depoimentos UGC reais são mais eficazes do que produção profissional aqui.'),

  h2('⑪ COMO TORNAR MAIS NATIVO', C.blue),
  bullet('Facebook: publicar como post pessoal sem logo de empresa. "Compartilhei isso com minha mãe depois de 3 anos vendo ela diminuir. Ela ficou brava. 6 meses depois me pediu desculpa. Aqui está o que eu mandei."'),
  bullet('TikTok: nenhuma produção. Celular estabilizado num objeto. Pessoa real olhando para câmera. Sem música nos primeiros 10 segundos — o silêncio é o pattern interrupt.'),
  bullet('Usar hashtags de nicho emocional, não de produto: #diabetes #identidade #autocuidado #aposentadoria em vez de #suplemento #glicemia #A1c.'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 6 — A SENTINELA CEGA
// ══════════════════════════════════════════════════════════════════════════════
const c6 = [
  ...conceptHeader(6,'A SENTINELA CEGA', ACCENTS[5]),

  h2('① NOME DO MECANISMO', C.green),
  pB('"Vigilância Metabólica Bloqueada" / "Sistema de Alerta Silenciado"'),
  p('O GLP-1 é posicionado como o "sistema de vigilância" do metabolismo — uma sentinela que monitora o açúcar 24 horas por dia e aciona a resposta de insulina. Quando a lama pancreática cobre esse sistema, a sentinela fica cega. O açúcar sobe sem que ninguém acione o alarme. O produto restaura a visão da sentinela.'),

  h2('② EXPLICAÇÃO PSICOLÓGICA DO MECANISMO', C.green),
  bullet('CONTROLE TÉCNICO: a metáfora de sistema de vigilância apela ao desejo de controle e monitoramento — especialmente poderoso para o público que já faz automonitoramento de glicemia (que é toda a audiência-alvo).'),
  bullet('INVERSÃO SURPREENDENTE: "seu corpo já tem um sistema de vigilância do açúcar — ele foi desativado" é mais surpreendente do que "você precisa de um novo tratamento". Transforma o problema de "falta de solução" para "solução existente que foi bloqueada".'),
  bullet('TECNOLOGIA COMO FAMILIARIDADE: câmeras de segurança, alertas e monitoramento são conceitos familiares para o público 50+ — não ameaçadores como termos médicos complexos.'),
  bullet('REATIVAÇÃO COMO PROMESSA: "reativar" é mais poderoso do que "tratar" — implica que o sistema original estava correto e apenas precisa ser desbloqueado. Remove a resignação.'),

  h2('③ NOME CHICLETE DA OFERTA', C.green),
  pR([{text:'Nome principal: ',bold:true},{text:'"Protocolo de Reativação" ou "Sistema Sentinela" ou "GLP-1 Reativado"'}]),
  pR([{text:'Variações: ',bold:true},{text:'"Alarme Metabólico" · "Fórmula Vigilância" · "Protocolo 24h" · "Código de Reativação"'}]),
  p('"Protocolo de Reativação" é preciso e aspiracional — soa como algo que um engenheiro faria em um sistema complexo que parou de funcionar. É técnico sem ser intimidante.'),

  h2('④ HOOK PRINCIPAL', C.green),
  hook('Dentro do seu pâncreas existe uma sentinela que monitora seu açúcar 24 horas por dia. Ela ficou cega. É por isso que seu açúcar sobe sem aviso. É por isso que você acorda às 3h da manhã com 400 de glicemia. Ninguém está vigiando. Aqui está o que reactiva essa sentinela.', ACCENTS[5]),

  h2('⑤ VARIAÇÕES DE HOOKS', C.green),
  p('Facebook:'),
  hookVar('Seu corpo tem um sistema de alarme para o açúcar. Ele foi silenciado pela lama tóxica que se acumula no pâncreas. É por isso que seu açúcar sobe sem que o seu organismo reaja a tempo. Aqui está o código de reativação.', C.green),
  hookVar('97% dos diabéticos tipo 2 têm o mesmo problema: o alarme que deveria controlar o açúcar foi silenciado. 4 ingredientes naturais reativam esse sistema em 21 dias. Aqui está como.', C.green),
  p(''),
  p('TikTok:'),
  hookVar('Seu pâncreas tem uma câmera de segurança 🎥 — e ela está com estática', C.green),
  hookVar('O alarme do açúcar no seu corpo foi silenciado. É por isso que o metformin não resolve 🔕', C.green),
  hookVar('97% dos diabéticos têm o sistema de vigilância do açúcar desativado. Aqui está o código de reativação ⚡', C.green),
  p(''),
  p('Telegram:'),
  hookVar('🔔 O sistema que deveria alertar seu pâncreas sobre o açúcar foi silenciado por depósitos tóxicos. Isso explica por que a glicemia sobe sem aviso e por que metformin nunca resolve o problema real. O protocolo de reativação está aqui:', C.blue),

  h2('⑥ IDEIA VISUAL DO CRIATIVO', C.green),
  tbl([
    ['Plataforma','Formato','Descrição Visual'],
    ['Facebook','Animação simples',
      'Interface de câmera de segurança: tela com estática (GLP-1 bloqueado). Alerta vermelho: "SINAL PERDIDO". Protocolo é aplicado. Tela volta nítida. Alerta verde: "SISTEMA ATIVO". Gráfico de glicemia estabilizando. 15 segundos. Parece interface de app de saúde real.'],
    ['TikTok','Split screen',
      'Esquerda: câmera de segurança real com estática + "SEM SINAL". Direita: câmera com imagem nítida + "MONITORAMENTO ATIVO". Legenda: "é assim que o GLP-1 funciona no seu pâncreas antes e depois de 21 dias no protocolo". Áudio: som de câmera de segurança ligando.'],
    ['Facebook','Carrossel infográfico',
      'Frame 1: diagrama simples — câmera de segurança (GLP-1) → sinal bloqueado (lama) → açúcar subindo sem controle. Frame 2: protocolo dissolvendo lama → câmera voltando → sinal restabelecido → insulina respondendo. Frame 3: resultados do trial. CTA.'],
  ],[1800,1800,5760], C.green),

  h2('⑦ EMOÇÃO PRINCIPAL', C.green),
  pR([{text:'Revelação Técnica → Urgência de Controle → Esperança de Reativação',bold:true,color:C.green}]),
  p('O público que faz automonitoramento de glicemia (toda a audiência-alvo) já está em modo de "vigilância" sobre sua saúde. Mostrar que existe um sistema interno de vigilância que foi silenciado valida o instinto de monitoramento e cria urgência técnica — "preciso reativar o sistema, não só monitorar externamente."'),

  h2('⑧ CTA IDEAL', C.green),
  tbl([
    ['Plataforma','CTA'],
    ['Facebook','Ver o Protocolo de Reativação Completo →'],
    ['TikTok','Código de reativação no link da bio ⚡'],
    ['Telegram','Protocolo de Reativação: [link]'],
  ],[2500,6860], C.green),

  h2('⑨ HEADLINE DA VSL', C.green),
  hook('"O Sistema de Vigilância do Açúcar que Existe no Seu Corpo Foi Silenciado — e Aqui Está o Protocolo de 4 Ingredientes que o Reactiva em 21 Dias"', ACCENTS[5]),

  h2('⑩ ÂNGULO DE ESCALA', C.green),
  p('Escala para nichos de tecnologia + saúde — audiências tech-savvy que usam wearables, CGMs (continuous glucose monitors) e apps de saúde. A metáfora de "sistema de vigilância" ressoa fortemente com quem já usa tecnologia de monitoramento. Criar variações que referenciam CGMs diretamente: "Seu CGM mostra o que acontece. O Protocolo de Reativação muda o que acontece."'),

  h2('⑪ COMO TORNAR MAIS NATIVO', C.green),
  bullet('TikTok: começar com imagem real de câmera de segurança com estática — parece um vídeo de problema técnico doméstico que vira revelação médica.'),
  bullet('Facebook: formato de "descoberta pessoal" — "Trabalhei 30 anos com sistemas de segurança. Quando me diagnosticaram com diabetes, percebi que era exatamente o mesmo problema: o sensor estava com defeito, não o sistema."'),
  bullet('Usar app de monitoramento de glicemia real na tela (com números altos) como ponto de partida visual — parece post de alguém compartilhando seu próprio exame.'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  CONCEITO 7 — O GÊMEO SAUDÁVEL
// ══════════════════════════════════════════════════════════════════════════════
const c7 = [
  ...conceptHeader(7,'O GÊMEO SAUDÁVEL', ACCENTS[6]),

  h2('① NOME DO MECANISMO', C.gold),
  pB('"DNA Dividido" / "O Paradoxo Genético"'),
  p('O mecanismo usa o twin study como prova irrefutável de que o diabetes não é genético — e que portanto pode ser revertido. A descoberta central: o que separa um gêmeo diabético do gêmeo saudável é um único hormônio (GLP-1), em níveis 312% diferentes. Mesmos genes. Resultado oposto. A diferença é tratável.'),

  h2('② EXPLICAÇÃO PSICOLÓGICA DO MECANISMO', C.gold),
  bullet('DESTRUIÇÃO DA RESIGNAÇÃO MAIS COMUM: "é hereditário na minha família" é a crença de resignação #1 do diabético tipo 2. O twin study a destrói de forma científica e irrefutável — se gêmeos idênticos têm resultados opostos, a genética não é destino.'),
  bullet('INJUSTIÇA COMO MOTIVADOR: "por que ele pode e eu não?" não é inveja — é a pergunta que precede a ação. A resposta ("porque seu GLP-1 está 312% mais baixo — e isso muda") transforma injustiça em possibilidade.'),
  bullet('ESPELHO GENÉTICO COMO PROXY PESSOAL: qualquer pessoa com familiar diabético pensa "isso pode ser eu." O gêmeo saudável é o "você que poderia ter sido" — e essa versão ainda é alcançável.'),
  bullet('CIÊNCIA EUROPEIA COMO CREDIBILIDADE: Lund University + Sweden = percepção de rigor científico neutro, sem influência de Big Pharma americana.'),

  h2('③ NOME CHICLETE DA OFERTA', C.gold),
  pR([{text:'Nome principal: ',bold:true},{text:'"Protocolo do Gêmeo Saudável" ou "DNA do Açúcar Perfeito"'}]),
  pR([{text:'Variações: ',bold:true},{text:'"A Chave Genética" · "Protocolo 312" · "Fator de Divergência" · "Código do Gêmeo"'}]),
  p('"Protocolo 312" é intrigante e específico — 312% é o número real do estudo de gêmeos (diferença de GLP-1). Um nome baseado em dado real que parece um código secreto.'),

  h2('④ HOOK PRINCIPAL', C.gold),
  hook('74 pares de gêmeos idênticos. Mesmo DNA. Mesmos pais. Mesma comida crescendo. Em todos os pares: um tem diabetes, o outro é completamente saudável. O que os separa é um único hormônio — em níveis 312% diferentes. Se um pode ter açúcar perfeito com o mesmo DNA, o outro também pode. Aqui está como cruzar essa diferença.', ACCENTS[6]),

  h2('⑤ VARIAÇÕES DE HOOKS', C.gold),
  p('Facebook (paradoxo científico):'),
  hookVar('Seu gêmeo idêntico não teria diabetes. Mesmo DNA que você — mas GLP-1 312% mais alto. A ciência prova: diabetes não é genético. É hormonal. E hormônios mudam. Aqui está como.', C.gold),
  hookVar('Se o diabetes fosse genético, gêmeos idênticos sempre teriam o mesmo resultado. Eles não têm. Em 74 pares estudados: um diabético, um saudável. Mesmo DNA. Aqui está o que os separa — e como mudar de lado.', C.gold),
  p(''),
  p('TikTok (paradoxo visual + choque):'),
  hookVar('Gêmeos idênticos: mesmo DNA, resultados OPOSTOS no diabetes 🧬 (estudo de 74 pares)', C.gold),
  hookVar('Isso prova que diabetes NÃO é genético 🤯 (e o que realmente causa)', C.gold),
  hookVar('Por que um gêmeo tem diabetes e o outro não? A resposta é um hormônio. E você pode mudá-lo.', C.gold),
  p(''),
  p('Native Ads / Advertorial:'),
  hookVar('PESQUISA: Gêmeos Idênticos com Diabetes em Apenas Um Dos Dois Derruba Teoria Genética — e Revela a Real Causa do Diabetes Tipo 2', C.gold),

  h2('⑥ IDEIA VISUAL DO CRIATIVO', C.gold),
  tbl([
    ['Plataforma','Formato','Descrição Visual'],
    ['Facebook','Split screen dramático',
      'Dois rostos quase idênticos (atores que se parecem). Esquerda: cansado, exames na mesa, medidor de glicemia. Direita: energético, mesa de jantar, sorrindo. Texto central: "MESMO DNA" → "GLP-1: 312% diferente". Dissolve para tabela comparativa de A1c. CTA.'],
    ['TikTok','Talking head + texto dinâmico',
      'Pessoa olhando para câmera. "Meu irmão gêmeo nunca teve diabetes na vida. Mesmo DNA que eu. Eu tomei metformin por 9 anos. Pesquisadores finalmente explicaram por que." [pausa dramática] "É um hormônio. E eu aprendi como mudar o meu." Texto na tela: GLP-1: 312% diferente.'],
    ['Native / Advertorial','Formato jornalístico',
      'Layout de artigo científico com foto de gêmeos + headline: "Estudo com 74 Pares de Gêmeos Idênticos Prova: Diabetes Não É Genético". Texto curto, link para VSL. Não parece anúncio — parece artigo compartilhado.'],
  ],[1800,1800,5760], C.gold),

  h2('⑦ EMOÇÃO PRINCIPAL', C.gold),
  pR([{text:'Injustiça → Curiosidade Científica → Esperança ("se ele pode, eu posso")',bold:true,color:C.gold}]),
  p('A injustiça do "mesmo DNA, resultado diferente" cria uma emoção única: não é medo, não é raiva — é incompreensão que exige resolução. O cérebro não tolera o paradoxo sem buscar a resposta. A resposta ("é o GLP-1") traz alívio imediato. E a solução ("você pode mudar isso") converte essa emoção em ação.'),

  h2('⑧ CTA IDEAL', C.gold),
  tbl([
    ['Plataforma','CTA'],
    ['Facebook','Descobrir o Que Nos Separa →'],
    ['TikTok','Protocolo 312: link na bio 🧬'],
    ['Native Ad','Leia o estudo completo — e o protocolo →'],
  ],[2500,6860], C.gold),

  h2('⑨ HEADLINE DA VSL', C.gold),
  hook('"O Estudo com 74 Pares de Gêmeos Idênticos que Provou: Diabetes Não É Genético — é Hormonal. E Aqui Está o Protocolo que Muda o Hormônio que Separa os Dois"', ACCENTS[6]),

  h2('⑩ ÂNGULO DE ESCALA', C.gold),
  p('Este é o conceito com maior potencial em native ads e advertorial (Taboola / Outbrain) — o formato jornalístico do paradoxo dos gêmeos é naturalmente clicável sem parecer propaganda. Escala para qualquer estudo científico paradoxal. Para Glycosept: criar série "A Ciência Que Contradiz Tudo" com o twin study como carro-chefe e outros estudos do VSL como sequência.'),

  h2('⑪ COMO TORNAR MAIS NATIVO', C.gold),
  bullet('Native: usar visual de artigo científico real (existe o estudo de Lund University — é real e verificável). Caption: "Pesquisadores da Suécia publicaram algo que contradiz 40 anos de medicina do diabetes."'),
  bullet('Facebook: post como se fosse um artigo compartilhado por um amigo. Texto: "Alguém me mandou isso hoje. Não consigo parar de pensar. Se for verdade, muda tudo sobre como entendemos o diabetes."'),
  bullet('TikTok: duet com um vídeo real de notícia sobre genética e diabetes — responder ao vídeo com a revelação do twin study. Parece reação, não anúncio.'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  SÍNTESE ESTRATÉGICA
// ══════════════════════════════════════════════════════════════════════════════
const synthesis = [
  pb(),
  h1('SÍNTESE ESTRATÉGICA — COMO ORQUESTRAR OS 7 CONCEITOS'),
  divider(C.navy),

  h2('MAPA DE TEMPERATURA DE TRÁFEGO'),
  tbl([
    ['Temperatura','Conceitos Recomendados','Motivo'],
    ['🥶 Cold (Zero Awareness)',    '①Código do Urso · ④Lama Doce · ⑦Gêmeo Saudável',  'Abre com curiosidade/espanto sem exigir conhecimento prévio do produto'],
    ['🌡 Morno (Retargeting)',      '②Interruptor Apagado · ⑥Sentinela Cega',           'Mecanismo técnico para quem já tem interesse e quer entender melhor'],
    ['🔥 Quente (Conversão)',       '③Hormônio Roubado · ⑤Espelho de 2 Anos',           'Atinge Deeper Core — para quem está perto de decidir mas ainda não converteu'],
  ],[2200,3800,3360], C.navy),

  p(''),
  h2('SEQUÊNCIA DE FUNIL RECOMENDADA'),
  tbl([
    ['Etapa','Conceito','Plataforma','Objetivo'],
    ['1 — Interrupção',  'Urso (TikTok 15s) ou Lama Doce (demo)',  'TikTok · Facebook cold', 'Parar o scroll. Zero menção a produto.'],
    ['2 — Educação',     'Gêmeo Saudável (native/advertorial)',      'Taboola · Outbrain',     'Construir crença no mecanismo.'],
    ['3 — Indignação',   'Hormônio Roubado (Facebook)',              'Facebook retargeting',   'Ativar raiva contra Big Pharma.'],
    ['4 — Identidade',   'Espelho de 2 Anos (Facebook/Telegram)',    'Facebook · Telegram',    'Tocar o Deeper Core antes da oferta.'],
    ['5 — Conversão',    'VSL completo (Glycosept)',                 'Landing page',           'Fechar com prova + urgência real.'],
  ],[1600,2800,2200,2760], C.navy),

  p(''),
  h2('REGRA DE OURO: O CONCEITO VIRAL'),
  p('O conceito com maior potencial de fenômeno viral é a combinação de:'),
  bullet('Hook: "Urso come mel puro e nunca fica diabético" (espanto biológico)'),
  bullet('Demo: dissolução da lama em funil de vidro (TikTok satisfying content)'),
  bullet('Climax: "97% reverteram o A1c em 12 semanas — uma cápsula por manhã"'),
  p('Esses três elementos juntos, em 30 segundos, têm o DNA de um vídeo que é compartilhado organicamente — parece conteúdo de ciência/natureza, não propaganda. É o ponto de partida para qualquer campanha que queira escalar além dos públicos diabéticos óbvios.'),

  p(''),
  h2('PLATAFORMAS E FORMATOS PRIORITÁRIOS'),
  tbl([
    ['Plataforma','Formato Ideal','Conceitos Prioritários','Métrica Foco'],
    ['Facebook Feed',   'Carrossel + Vídeo 60s', '③⑤⑦',           'CPM · CTR · Comment Rate'],
    ['Facebook Reels',  'Vídeo vertical 15-30s', '①②④',           'Hook Rate (3s) · View Through'],
    ['TikTok',          'Vídeo vertical 15-30s', '①④⑥',           'Watch Time · Share Rate · Comment'],
    ['YouTube Pre-roll','Vídeo 30s não pulável', '①③⑦',           'View Through Rate · CTR'],
    ['Telegram',        'Texto longo + link',    '③⑤',            'Click Rate · Forward Rate'],
    ['Native (Taboola)','Artigo + imagem',       '⑦ (Twin Study)','CTR · Scroll Depth · VSL Start'],
  ],[1600,2000,2000,3760], C.navy),
];

// ══════════════════════════════════════════════════════════════════════════════
//  ASSEMBLE
// ══════════════════════════════════════════════════════════════════════════════
const allChildren = [
  ...cover, ...c1, ...c2, ...c3, ...c4, ...c5, ...c6, ...c7, ...synthesis
];

const doc = new Document({
  numbering:{ config:[
    { reference:'bullets', levels:[{ level:0, format:LevelFormat.BULLET, text:'•',
        alignment:AlignmentType.LEFT, style:{paragraph:{indent:{left:720,hanging:360}}}}]},
    { reference:'numbers', levels:[{ level:0, format:LevelFormat.DECIMAL, text:'%1.',
        alignment:AlignmentType.LEFT, style:{paragraph:{indent:{left:720,hanging:360}}}}]},
  ]},
  styles:{
    default:{ document:{ run:{ font:'Arial', size:22 }}},
    paragraphStyles:[
      { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{size:38,bold:true,font:'Arial',color:C.navy}, paragraph:{spacing:{before:400,after:200},outlineLevel:0}},
      { id:'Heading2', name:'Heading 2', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{size:28,bold:true,font:'Arial',color:C.blue}, paragraph:{spacing:{before:300,after:160},outlineLevel:1}},
      { id:'Heading3', name:'Heading 3', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{size:24,bold:true,font:'Arial',color:C.navy}, paragraph:{spacing:{before:220,after:100},outlineLevel:2}},
    ]
  },
  sections:[{
    properties:{ page:{ size:{width:12240,height:15840}, margin:{top:1440,right:1440,bottom:1440,left:1440}}},
    headers:{ default: new Header({ children:[new Paragraph({
      alignment:AlignmentType.RIGHT,
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.orange,space:6}},
      children:[new TextRun({text:'7 CONCEITOS DE ADS — GLYCOSEPT | Atlas ADS',size:18,font:'Arial',color:C.gray})]
    })]})},
    footers:{ default: new Footer({ children:[new Paragraph({
      alignment:AlignmentType.CENTER,
      children:[
        new TextRun({text:'Pág. ',size:18,font:'Arial',color:C.gray}),
        new TextRun({children:[PageNumber.CURRENT],size:18,font:'Arial',color:C.gray}),
        new TextRun({text:' | Atlas ADS — Confidencial',size:18,font:'Arial',color:C.gray})
      ]
    })]})},
    children: allChildren
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('/sessions/loving-dazzling-knuth/mnt/outputs/7_Conceitos_Ads_Glycosept.docx', buf);
  console.log('Gerado com sucesso!');
}).catch(e=>{ console.error(e); process.exit(1); });
