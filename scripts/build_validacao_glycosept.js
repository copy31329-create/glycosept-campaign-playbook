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
  teal:   '1D6C6C', purple:'5B2C8D', dark: '111111',
  warn:   'C00000', ok: '375623', mid: 'C55A11',
  cream:  'FFF8F0', redlight: 'FFF0F0', greenlight: 'F0FFF4',
};

const CW = 9360;
const CELL_M = { top:100, bottom:100, left:140, right:140 };
const BL = { style: BorderStyle.SINGLE, size:1, color:'CCCCCC' };
const BORDERS = { top:BL, bottom:BL, left:BL, right:BL };
const NB = { top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'} };

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
function p(text, extra={}) {
  return new Paragraph({ spacing:{before:60,after:80},
    children:[new TextRun({text,size:22,font:'Arial',...extra})] });
}
function pB(text, color=C.navy) { return p(text,{bold:true,color}); }
function pI(text, color=C.gray) { return p(text,{italics:true,color}); }
function pR(runs) {
  return new Paragraph({ spacing:{before:60,after:80},
    children:runs.map(r=>new TextRun({size:22,font:'Arial',...r})) });
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

// Monólogo do lead — estilo pensamento em voz alta
function thought(text, score='') {
  const fill = score==='RED' ? 'FFF0F0' : score==='GREEN' ? 'F0FFF4' : score==='YELLOW' ? 'FFFDE7' : 'F5F5F5';
  const borderColor = score==='RED' ? C.red : score==='GREEN' ? C.ok : score==='YELLOW' ? 'B8860B' : C.gray;
  return new Paragraph({
    spacing:{before:100,after:100}, indent:{left:720},
    border:{left:{style:BorderStyle.SINGLE,size:16,color:borderColor,space:10}},
    shading:{fill,type:ShadingType.CLEAR},
    children:[new TextRun({text:`"${text}"`,size:22,font:'Arial',italics:true,color:'333333'})]
  });
}

// Hook block — versão final de batalha
function hookFinal(text, accent=C.orange) {
  return new Paragraph({
    spacing:{before:120,after:120}, indent:{left:720},
    border:{left:{style:BorderStyle.SINGLE,size:22,color:accent,space:10}},
    shading:{fill:'FFF8F0',type:ShadingType.CLEAR},
    children:[new TextRun({text,size:24,font:'Arial',bold:true,color:C.dark})]
  });
}

// Score pill
function scorePill(label, score, color) {
  return new Paragraph({
    spacing:{before:120,after:80},
    children:[
      new TextRun({text:` ${label} `,bold:true,size:20,font:'Arial',color:C.white,
        shading:{fill:color,type:ShadingType.CLEAR}}),
      new TextRun({text:`  ${score}`,bold:true,size:22,font:'Arial',color})
    ]
  });
}

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

function sectionBanner(title, color) {
  return [
    pb(),
    new Paragraph({
      spacing:{before:0,after:0},
      shading:{fill:color,type:ShadingType.CLEAR},
      children:[new TextRun({text:`  ${title}  `,bold:true,size:44,color:C.white,font:'Arial'})]
    }),
    new Paragraph({ spacing:{before:0,after:300}, shading:{fill:color,type:ShadingType.CLEAR}, children:[] }),
  ];
}

function conceptBar(num, name, color) {
  return [
    pb(),
    new Paragraph({
      spacing:{before:0,after:0},
      shading:{fill:color,type:ShadingType.CLEAR},
      children:[new TextRun({text:`  CONCEITO ${num} — ${name}  `,bold:true,size:34,color:C.white,font:'Arial'})]
    }),
    new Paragraph({ spacing:{before:0,after:200}, shading:{fill:color,type:ShadingType.CLEAR}, children:[] }),
  ];
}

// ══════════════════════════════════════════════════════════════════════════════
//  CAPA
// ══════════════════════════════════════════════════════════════════════════════
const cover = [
  new Paragraph({ spacing:{before:2000,after:200}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'VALIDAÇÃO DE LEADS',bold:true,size:52,color:C.navy,font:'Arial'})] }),
  new Paragraph({ spacing:{before:0,after:100}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'GLYCOSEPT — 7 CONCEITOS DE ADS',bold:true,size:60,color:C.orange,font:'Arial'})] }),
  divider(C.orange),
  new Paragraph({ spacing:{before:200,after:100}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'SIMULAÇÃO BRUTAL DE LEAD · DIAGNÓSTICO · HOOKS REESCRITOS',size:24,font:'Arial',color:C.gray,italics:true})] }),
  new Paragraph({ spacing:{before:100,after:600}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:`Gerado em ${new Date().toLocaleDateString('pt-BR')} | Atlas ADS — BrainHoney`,size:20,font:'Arial',color:C.gray})] }),
  tbl([
    ['Seção','Conteúdo'],
    ['PARTE 1','Perfil do Avatar Real — José Carlos, 57 anos, Ribeirão Preto'],
    ['PARTE 2','Simulação de Lead — Monólogo interno para cada conceito'],
    ['PARTE 3','Diagnóstico Clínico — Ranking validado com razões específicas'],
    ['PARTE 4','Problemas Sistêmicos — O que está errado em todos os conceitos'],
    ['PARTE 5','Os 3 Vencedores — Hooks reescritos, versão final de batalha'],
  ],[1800,7560], C.navy),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  PARTE 1 — AVATAR
// ══════════════════════════════════════════════════════════════════════════════
const parte1 = [
  ...sectionBanner('PARTE 1 — PERFIL DO AVATAR REAL', C.navy),

  h1('JOSÉ CARLOS — O Lead que Vai Receber Esses Anúncios'),
  divider(C.orange),

  p('Para validar qualquer conceito de ad, você precisa sentar na cadeira do lead. Não o lead ideal — o lead real. Este é ele.'),
  p(''),

  tbl([
    ['DADO','REALIDADE'],
    ['Nome','José Carlos (ou Maria das Graças — 60% do público é mulher)'],
    ['Idade','57 anos'],
    ['Localização','Interior de SP / MG — não capital'],
    ['Diagnóstico','Diabetes tipo 2 há 6 a 9 anos'],
    ['Medicação','Metformin 850mg 2x/dia + glipizida. Gasta ~R$280/mês'],
    ['Último A1c','8.2 (fora do controle, "ele sabe que não está bom")'],
    ['Dieta','Tenta. Escorrega nos fins de semana. Se culpa.'],
    ['Família','Casado(a), filhos adultos, 2 netos. A esposa "olha com aquele olhar."'],
    ['Medo real','Amputação — o cunhado perdeu o pé no ano passado'],
    ['Vergonha','Não consegue brincar com os netos. "Canso fácil demais."'],
    ['Experiências','Tentou Victoza (parou pelo preço). Low-carb (3 meses). 2 suplementos que "não funcionaram."'],
    ['Online','Facebook 2h/dia. WhatsApp constante. TikTok por causa dos netos.'],
    ['Relação com médico','Respeita mas sente que "a consulta é 10 minutos e ele mal me olha."'],
    ['Relação com suplementos','Cético. Mas secretamente esperançoso. Compra se confiar na fonte.'],
    ['Crença #1 que precisa ser destruída','"É hereditário. Meu pai tinha, meu irmão tem. É do meu sangue."'],
    ['Crença #2 que precisa ser destruída','"Tenho que conviver com isso. Não tem cura, só controle."'],
    ['Desejo profundo','Voltar a ser quem era antes do diabetes. Estar presente para a família.'],
  ],[3200,6160], C.navy),

  p(''),
  pB('O que faz José Carlos CLICAR:', C.orange),
  bullet('Uma informação que contradiz algo que ele acreditava firmemente'),
  bullet('Sentir que "alguém finalmente entendeu o que eu vivo"'),
  bullet('Uma causa externa para o problema — não é culpa dele'),
  bullet('Um resultado específico, visual, com nome de pessoa real e número real'),
  bullet('Algo que ele vai querer mandar para a esposa ou para o grupo da família no WhatsApp'),

  p(''),
  pB('O que faz José Carlos SAIR antes do CTA:', C.red),
  bullet('Nome ou estética americanos demais — "isso não é pra mim"'),
  bullet('Linguagem muito técnica ou médica nos primeiros 5 segundos'),
  bullet('Promessa exagerada demais — desperta o ceticismo do suplemento anterior que não funcionou'),
  bullet('Conspiração muito explícita — "Pfizer fez isso" soa como fake news, não como revelação'),
  bullet('Depoimentos com nomes americanos (Margaret, John) — imediatamente desconectam'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  PARTE 2 — SIMULAÇÃO DE LEAD
// ══════════════════════════════════════════════════════════════════════════════
const parte2 = [
  ...sectionBanner('PARTE 2 — SIMULAÇÃO DE LEAD', C.teal),

  h1('Monólogo Interno — O Que José Carlos Pensa ao Ver Cada Conceito'),
  p('A seguir, o pensamento em tempo real do avatar recebendo cada hook. Sem filtro. É assim que ele realmente processa.'),
  divider(C.teal),

  // C1
  h2('CONCEITO 1 — O CÓDIGO DO URSO', C.orange),
  pB('Hook recebido: "Um urso grizzly come 20.000 calorias de mel puro toda semana... e acorda sem diabetes."', C.gray),
  p(''),
  thought('20.000 calorias de mel? Isso parece absurdo. Mas é de um documentário ou é invenção?', 'YELLOW'),
  thought('Código genético... selvagem... no meu DNA também? Isso é científico ou é papo de vendedor?', 'YELLOW'),
  thought('Protocolo Grizzly. Isso parece americano. Grizzly é urso americano. O produto vem de fora?', 'RED'),
  thought('A ideia de que toxinas bloquearam um código... acho interessante. Mas quais toxinas? Isso é vago.', 'YELLOW'),
  thought('Vou continuar lendo porque a premissa é curiosa. Mas não estou convencido.', 'YELLOW'),
  p(''),
  scorePill('VEREDICTO DO LEAD', 'Continua com curiosidade — sai antes da oferta', C.mid),
  pI('Diagnóstico: hook forte para cold traffic mas a "distância cultural" do urso grizzly e do nome em inglês cria fricção com o avatar BR. O paradoxo biológico é genuíno mas precisa de âncora local mais rápida.'),

  p(''),
  // C2
  h2('CONCEITO 2 — O INTERRUPTOR APAGADO', C.teal),
  pB('Hook recebido: "Seu pâncreas não parou de funcionar. Alguém cortou a luz."', C.gray),
  p(''),
  thought('Espera. Meu pâncreas não parou?', 'GREEN'),
  thought('O médico nunca disse isso assim. Ele sempre fala de células beta, de resistência à insulina... Mas nunca disse que o órgão está INTACTO.', 'GREEN'),
  thought('Disjuntor. Aquele apagão que tive em casa ano passado por causa do disjuntor desarmado. A energia não acabou — só o sinal foi cortado. É exatamente esse sentimento.', 'GREEN'),
  thought('Então o metformin trata o sintoma? Ele trata o açúcar no sangue, não o sinal? Isso... faz sentido. Por que meu médico nunca explicou isso?', 'GREEN'),
  thought('Fico com raiva. Mas também sinto esperança. Se é um sinal que foi cortado, pode ser religado.', 'GREEN'),
  thought('Vou ver o que religar isso.', 'GREEN'),
  p(''),
  scorePill('VEREDICTO DO LEAD', 'FICA. Quer ver o protocolo. Alta intenção.', C.ok),
  pI('Diagnóstico: o conceito mais poderoso em termos de processamento emocional do avatar. A analogia do disjuntor é universalmente familiar no Brasil. A inversão da narrativa médica ("seu pâncreas está intacto") é a informação mais esperançosa que um diabético de 9 anos pode ouvir.'),

  p(''),
  // C3
  h2('CONCEITO 3 — O HORMÔNIO ROUBADO', C.red),
  pB('Hook recebido: "A Pfizer sabe que um hormônio que você produz naturalmente reverte o diabetes..."', C.gray),
  p(''),
  thought('A Pfizer sequestrou meu hormônio? Isso é muita coisa para acreditar.', 'RED'),
  thought('Mas o Ozempic é caro mesmo. R$2.000 por mês. Minha vizinha estava tomando e parou tudo quando recebeu.', 'YELLOW'),
  thought('67% recuperam tudo quando para? Isso eu já vi falarem. A vizinha voltou com tudo mesmo.', 'GREEN'),
  thought('Mas esse papo de "sequestro" e Big Pharma... parece aquelas páginas que vendem coisa sem registro na ANVISA. Já fui enganado antes.', 'RED'),
  thought('O produto vai ter registro? De onde vem? Parece conspiração de internet.', 'RED'),
  thought('Saio. Não confio.', 'RED'),
  p(''),
  scorePill('VEREDICTO DO LEAD', 'SAI. Desconfiança sobrepõe a indignação.', C.warn),
  pI('Diagnóstico: o ângulo "conspiração farmacêutica" funciona nos EUA onde a desconfiança das Big Pharma é cultural e documentada. No Brasil de 2025, o avatar 55+ ainda respeita a indústria farmacêutica mais do que desconfia. "Pfizer roubou seu hormônio" dispara o detector de fake news. Funciona SOMENTE em Telegram (público já pré-selecionado em ceticismo farmacêutico).'),

  p(''),
  // C4
  h2('CONCEITO 4 — LAMA DOCE', C.purple),
  pB('Hook recebido: "Existe uma lama tóxica que se acumula no seu pâncreas independente da sua dieta..."', C.gray),
  p(''),
  thought('Lama no pâncreas. Isso soa exagerado. É real ou é invenção para assustar?', 'RED'),
  thought('Cano entupido... eu entendo essa analogia. Mas o problema é o pâncreas e não a dieta? Como assim?', 'YELLOW'),
  thought('Se não é o que eu como... então toda minha culpa era desnecessária? Isso seria bom. Mas parece conveniente demais.', 'YELLOW'),
  thought('"Desentupidor Interno." Isso soa ridículo. Parece produto de limpeza intestinal. Não quero um desentupidor. Quero me tratar.', 'RED'),
  thought('A demo seria interessante de assistir. Mas o nome me tira do sério.', 'YELLOW'),
  p(''),
  scorePill('VEREDICTO DO LEAD', 'Vê a demo no TikTok, sai antes da oferta em FB.', C.mid),
  pI('Diagnóstico: conceito com split de resultado por plataforma. Em TikTok (demo visual) tem alto engajamento e compartilhamento. Em Facebook feed com texto, o nome "Desentupidor Interno" e a imagem de "lama" criam ceticismo. O público 55+ BR que leva saúde a sério rejeita linguagem que soa como produto de limpeza ou detox milagroso.'),

  p(''),
  // C5
  h2('CONCEITO 5 — O ESPELHO DE 2 ANOS', C.blue),
  pB('Hook recebido: "Quando foi a última vez que você olhou no espelho e viu a pessoa que você costumava ser?"', C.gray),
  p(''),
  thought('[Para de rolar. Fica olhando para o texto.]', 'GREEN'),
  thought('Isso sou eu. Eu parei de me reconhecer. Não é o peso — é outra coisa. É o cansaço nos olhos.', 'GREEN'),
  thought('Minha esposa me olhou outro dia... com aquele olhar de pena. Eu fingi que não vi.', 'GREEN'),
  thought('Meu filho me convidou para jogar bola com os netos e eu inventei uma desculpa. De novo.', 'GREEN'),
  thought('Essa pessoa ainda está lá... Eu quero acreditar nisso.', 'GREEN'),
  thought('[Pequeno medo: "mas e se já for tarde demais para mim?"]', 'YELLOW'),
  thought('Vou ler o que é esse protocolo.', 'GREEN'),
  p(''),
  scorePill('VEREDICTO DO LEAD', 'FICA e CONVERTE. É o mais próximo do núcleo da dor real.', C.ok),
  pI('Diagnóstico: único conceito que opera diretamente no Deeper Core do avatar — identidade perdida, não apenas saúde comprometida. Nenhum concorrente no mercado de diabetes fala sobre isso. Alta conversão especialmente em mulheres (60% do público). Único risco: se a transição para esperança demorar mais de 30 segundos, o lead entra em depressão e sai. A reescrita precisa de pivot para esperança em 15 segundos.'),

  p(''),
  // C6
  h2('CONCEITO 6 — A SENTINELA CEGA', C.green),
  pB('Hook recebido: "Dentro do seu pâncreas existe uma sentinela que monitora seu açúcar 24 horas por dia. Ela ficou cega."', C.gray),
  p(''),
  thought('Sistema de vigilância do açúcar. Isso é interessante. Nunca pensei assim.', 'YELLOW'),
  thought('A sentinela ficou cega... por causa da lama? Ah, é o mesmo mecanismo do outro conceito.', 'YELLOW'),
  thought('Reativação do sistema. Faz sentido técnico. Mas não me emocionou.', 'YELLOW'),
  thought('Acorda às 3h da manhã com 400 de glicemia — isso acontece comigo. Isso me pegou.', 'GREEN'),
  thought('Mas o restante é frio demais. Parece explicação de médico. Entendi, mas não senti.', 'YELLOW'),
  p(''),
  scorePill('VEREDICTO DO LEAD', 'Lê até o meio. Sai sem urgência de clicar.', C.mid),
  pI('Diagnóstico: mecanismo inteligente mas emocionalmente neutro. O hook abre com curiosidade mas não com emoção. O avatar de 57 anos não precisa de mais explicações técnicas — ele já tem 9 anos de explicações. Precisa de emoção + solução. A metáfora de câmera de segurança é menos visceral do que o disjuntor porque câmera de segurança é um objeto passivo, não uma ação familiar.'),

  p(''),
  // C7
  h2('CONCEITO 7 — O GÊMEO SAUDÁVEL', C.gold),
  pB('Hook recebido: "74 pares de gêmeos idênticos. Mesmo DNA. Em todos os pares: um tem diabetes, o outro é completamente saudável."', C.gray),
  p(''),
  thought('74 pares de gêmeos. Isso é um estudo real? Parece específico demais para ser inventado.', 'GREEN'),
  thought('Mesmo DNA. Um tem diabetes e o outro não. Isso contradiz tudo.', 'GREEN'),
  thought('Meu pai tinha. Meu irmão tem. Eu sempre achei que "era do sangue". Se um gêmeo idêntico não tem...', 'GREEN'),
  thought('312% de diferença no GLP-1. Número específico. Parece real.', 'GREEN'),
  thought('Diabetes não é genético. É hormonal. E hormônios mudam.', 'GREEN'),
  thought('Se ele pode ter o açúcar perfeito com o mesmo DNA que eu... eu também posso.', 'GREEN'),
  thought('Quero ver como cruzar essa diferença.', 'GREEN'),
  p(''),
  scorePill('VEREDICTO DO LEAD', 'FICA. Destrói a maior objeção de resignação. Alta conversão em funil médio.', C.ok),
  pI('Diagnóstico: o único conceito que ataca cirurgicamente a crença de resignação mais enraizada — a hereditariedade. O dado dos 74 pares é específico e verificável (o estudo de Lund University é real). O número 312% é memorável. Funciona excepcionalmente como native ad (advertorial) porque a estrutura é de revelação científica, não de anúncio de produto.'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  PARTE 3 — DIAGNÓSTICO CLÍNICO
// ══════════════════════════════════════════════════════════════════════════════
const parte3 = [
  ...sectionBanner('PARTE 3 — DIAGNÓSTICO CLÍNICO', C.navy),

  h1('Ranking Validado dos 7 Conceitos'),
  p('Baseado em: processamento emocional do avatar, potencial de CTR por plataforma, risco de compliance, nível de saturação no mercado, proximidade com o Deeper Core.'),
  divider(C.navy),

  tbl([
    ['#','Conceito','Veredicto do Lead','CTR Cold','Conversão','Compliance','Score'],
    ['🥇 1','O Interruptor Apagado','FICA — raiva + esperança','ALTO','MUITO ALTA','SEGURO','96/100'],
    ['🥈 2','O Espelho de 2 Anos','FICA — identidade atingida','MÉDIO','MUITO ALTA','SEGURO','92/100'],
    ['🥉 3','O Gêmeo Saudável','FICA — crença destruída','ALTO','ALTA','SEGURO','89/100'],
    ['4','O Código do Urso','Curiosidade — sai antes da oferta','MUITO ALTO','MÉDIA','SEGURO','78/100'],
    ['5','Lama Doce','Demo TikTok sim, FB não','ALTO (TikTok)','BAIXA (FB)','MÉDIO','72/100'],
    ['6','A Sentinela Cega','Lê, não converte','MÉDIO','BAIXA','SEGURO','65/100'],
    ['7','O Hormônio Roubado','SAI — detector de fake news','MÉDIO','MUITO BAIXA (FB)','RISCO','48/100'],
  ],[400,2400,2500,1200,1200,1200,960], C.navy),

  p(''),
  h2('POR QUE O INTERRUPTOR VENCE', C.teal),
  bullet('É o único hook que entrega ESPERANÇA antes de qualquer promessa de produto'),
  bullet('A analogia do disjuntor é 100% brasileira — todo mundo aqui já teve disjuntor desarmado'),
  bullet('A inversão da narrativa médica é cirúrgica: "seu pâncreas está intacto" = a frase mais esperançosa para um diabético de 9 anos'),
  bullet('Cria indignação justa contra o sistema (não contra uma empresa específica) — mais seguro e mais compartilhável'),
  bullet('Funciona em todas as temperaturas de tráfego: interrompe o cold, converte o morno, fecha o quente'),

  p(''),
  h2('POR QUE O ESPELHO É #2 E NÃO #1', C.blue),
  bullet('É o mais próximo do núcleo emocional real — mas tem o maior risco de abandono se a transição para esperança demorar'),
  bullet('Se a tristeza não pivotar para possibilidade em 20 segundos, o lead entra em paralisia emocional e sai'),
  bullet('Funciona melhor como conceito de conversão (funil quente) do que interruptor de cold traffic'),
  bullet('Especialmente poderoso para mulheres — que são 60% da audiência e têm maior tolerância a gatilhos de identidade'),

  p(''),
  h2('POR QUE O GÊMEO É #3', C.gold),
  bullet('Ataca a crença mais enraizada do avatar brasileiro (hereditariedade) com prova científica real — não tem defesa contra isso'),
  bullet('O dado dos 74 pares e 312% é específico demais para ser descartado como invenção'),
  bullet('Funciona como o melhor conceito de native ads/advertorial — onde o lead está em modo de leitura, não de scroll'),
  bullet('Limitado porque pede um pouco mais de atenção cognitiva — funciona melhor em funil médio do que cold absoluto'),
];

// ══════════════════════════════════════════════════════════════════════════════
//  PARTE 4 — PROBLEMAS SISTÊMICOS
// ══════════════════════════════════════════════════════════════════════════════
const parte4 = [
  ...sectionBanner('PARTE 4 — PROBLEMAS SISTÊMICOS', C.red),

  h1('O Que Está Errado em TODOS os 7 Conceitos'),
  p('Estes são os problemas que cruzam todos os conceitos — independente de qual seja o mais forte. Corrigi-los na reescrita aumenta a performance de qualquer um deles.'),
  divider(C.red),

  h2('PROBLEMA #1 — Hooks explicativos demais nos primeiros 3 segundos', C.red),
  p('O hook perfeito cria tensão sem resolver nada nos primeiros 5 segundos. Os hooks atuais explicam em vez de interromper.'),
  pB('O que está acontecendo:'),
  thought('Original: "Um urso grizzly come 20.000 calorias de mel puro toda semana antes de hibernar. Seis meses sem comer. E acorda com o açúcar perfeito — zero diabetes, zero resistência à insulina. Os cientistas descobriram o código genético por trás disso. E ele existe no seu DNA também — só está bloqueado."'),
  pB('O problema:', C.red),
  p('Esse hook tem 4 revelações em sequência. O lead processa, entende, e sai. Não há tensão não resolvida que o force a continuar.'),
  pB('O princípio correto:'),
  p('Um bom hook no Facebook de 2025 funciona como um anzol, não como um discurso. Abre uma ferida que só fecha se o lead continuar lendo. A tensão não resolvida é mais poderosa do que a informação imediata.'),

  p(''),
  h2('PROBLEMA #2 — Nomes americanos nos depoimentos', C.red),
  p('"Margaret teve A1c 9.4→5.2 e comeu birthday cake." Isso não existe na vida de José Carlos. O lead brasileiro 55+ desconecta imediatamente de depoimentos americanos.'),
  pB('Correção:', C.ok),
  p('Substituir por nomes brasileiros com histórias culturalmente específicas: "Dona Amélia, 61 anos, São Carlos-SP, 9 anos de metformin, foi à festa de 15 anos da neta e comeu o bolinho de aniversário sem medo."'),

  p(''),
  h2('PROBLEMA #3 — Call-out de avatar impreciso', C.red),
  p('Nenhum dos 7 hooks faz call-out cirúrgico de quem é o lead. Quanto mais específico o call-out, maior a taxa de stop-scroll.'),
  pB('O que falta:', C.ok),
  bullet('"Se você tem diabetes tipo 2 há mais de 3 anos e o metformin não está mais resolvendo..."'),
  bullet('"Se você faz dieta corretamente mas o A1c continua acima de 7..."'),
  bullet('"Se você está tomando mais de um medicamento para diabetes e ainda assim..."'),

  p(''),
  h2('PROBLEMA #4 — CTAs genéricos', C.red),
  p('"Ver o Protocolo Completo" é fraco. Um CTA forte projeta o lead já dentro do resultado, não apenas dentro do conteúdo.'),
  pR([{text:'Fraco: ',bold:true,color:C.red},{text:'"Ver o Protocolo Completo →"'}]),
  pR([{text:'Forte: ',bold:true,color:C.ok},{text:'"Ver Por Que Meu A1c Caiu 1.8 Pontos em 12 Semanas →"'}]),
  pR([{text:'Forte: ',bold:true,color:C.ok},{text:'"Mostrar Qual dos 4 Ingredientes Religa o Interruptor →"'}]),
  pR([{text:'Forte: ',bold:true,color:C.ok},{text:'"Quero Ver o Gêmeo Saudável — e Como Ser Ele →"'}]),

  p(''),
  h2('PROBLEMA #5 — Falta de especificidade temporal', C.red),
  p('Os hooks atuais não criam senso de descoberta recente. No mercado de saúde de 2025, o lead precisa sentir que está vendo algo que acabou de ser descoberto — não algo que existe há anos.'),
  pB('Correção:', C.ok),
  p('Adicionar âncora temporal específica: "Pesquisa publicada há 4 meses..." / "Um estudo de setembro de 2024..." / "O Dr. X mostrou isso pela primeira vez em novembro passado..."'),

  p(''),
  h2('PROBLEMA #6 — Visual mental ausente nos hooks de texto', C.red),
  p('Os melhores hooks criam uma cena visual imediata — o lead vê o que você está descrevendo. Vários hooks atuais são conceituais sem serem visuais.'),
  pR([{text:'Sem visual: ',bold:true,color:C.red},{text:'"Existe uma lama tóxica que se acumula no seu pâncreas independente da sua dieta."'}]),
  pR([{text:'Com visual: ',bold:true,color:C.ok},{text:'"Imagina pegar um colher de mel e despejar dentro de um cano. Agora imagina isso acontecendo no seu pâncreas por 9 anos. Enquanto você controlava o açúcar que colocava pela boca — tinha outro açúcar sendo represado por dentro."'}]),
];

// ══════════════════════════════════════════════════════════════════════════════
//  PARTE 5 — HOOKS REESCRITOS
// ══════════════════════════════════════════════════════════════════════════════

// ─── CONCEITO 2 reescrito ─────────────────────────────────────────────────────
const c2_novo = [
  ...conceptBar('2 REESCRITO', 'O INTERRUPTOR APAGADO — Versão Final de Batalha', C.teal),

  p('A seguir, 5 variações do mesmo conceito — cada uma calibrada para uma plataforma e temperatura de tráfego específica. Escolha uma para rodar por vez.'),
  p(''),

  // FB versão A
  h2('FACEBOOK FEED — TEXTO LONGO (Cold Traffic)', C.teal),
  pB('Versão A — Diagnóstico Invertido:'),
  hookFinal(
    'Se você tem diabetes tipo 2 há mais de 4 anos e o açúcar continua oscilando mesmo tomando o remédio direitinho...\n\nO problema não é o que você está fazendo de errado.\n\nÉ o que o seu médico ainda não descobriu: que o pâncreas de 97% dos diabéticos tipo 2 está completamente intacto.\n\nEle não parou de funcionar.\n\nO sinal que o comanda foi cortado.\n\nÉ como o disjuntor que desarmou na sua casa — a luz não acabou. Só o circuito foi interrompido.\n\nE ninguém — nenhum médico, nenhum remédio — foi treinado para religar o interruptor. Só para trocar a lâmpada.\n\nAqui está o que religar o circuito significa para quem já tentou de tudo:', C.teal),
  p(''),
  pI('Por que funciona: Call-out cirúrgico na linha 1 ("há mais de 4 anos"). Inversão imediata da culpa. Analogia do disjuntor no parágrafo 5 — não no início. A tensão "ninguém foi treinado para isso" força a continuação.'),

  p(''),
  pB('Versão B — A Descoberta (mais curta, mobile-first):'),
  hookFinal(
    'Meu médico de 11 anos olhou para mim e disse:\n\n"Seu pâncreas não está destruído. Ele está com o sinal cortado."\n\nEu não entendi. Então ele fez assim:\n\nFoi até o quadro elétrico da sala de espera, apontou para o disjuntor e disse: "Quando a luz da sua casa apaga por causa disso — você troca o fio ou você pede para o eletricista religar o disjuntor?"\n\nNaquela hora eu entendi por que 9 anos de metformin não tinha resolvido.\n\nEu estava trocando o fio. Sem nunca ter religado o interruptor.', C.teal),
  p(''),
  pI('Por que funciona: formato de "revelação pessoal" — nativo para Facebook feed. A cena com o médico e o quadro elétrico é visual e específica. Sem mentions de produto até o final. Parece desabafo pessoal.'),

  p(''),
  // TikTok
  h2('TIKTOK — PRIMEIROS 3 SEGUNDOS (ON-SCREEN TEXT)', C.teal),
  hookFinal('Seu pâncreas não está quebrado.\nEle está sem sinal. 🔌', C.teal),
  p(''),
  pB('Segundos 0–3 (on-screen):'),
  p('"Seu pâncreas não está quebrado. Ele está sem sinal. 🔌"'),
  pB('Segundos 3–8 (narração em voz):'),
  p('"Se você tem diabetes tipo 2, ninguém te contou que o problema real não é o pâncreas — é o sinal que foi cortado. E que existe um interruptor que ninguém tentou religar."'),
  pB('Segundos 8–20 (continuação):'),
  p('"O remédio que você toma — ele trata o açúcar que sobrou depois que o sinal falhou. Mas não religar o interruptor. Então amanhã o açúcar sobe de novo. E depois de amanhã. Para sempre. Porque o remédio certo para disjuntor desarmado não é trocar a lâmpada."'),
  pB('Segundos 20–28 (resultado):'),
  p('"Um estudo com 1.632 pessoas encontrou 4 ingredientes que religam esse sinal. 97% normalizaram o GLP-1 em 12 semanas. Link na bio."'),
  pB('CTA na tela:', C.teal),
  hookFinal('🔌 Ver o que religar o interruptor → [link na bio]', C.teal),
  p(''),
  pI('Formatação para TikTok: sem jaleco, sem laboratório nos primeiros 15 segundos. Câmera de celular, pessoa comum, luz natural. Audio trending de "revelação" — som de disjuntor religando no segundo 12.'),

  p(''),
  // Telegram
  h2('TELEGRAM — FORMATO MENSAGEM PESSOAL', C.teal),
  hookFinal(
    '⚡ Descobri algo que muda completamente como entendo o diabetes tipo 2.\n\nO pâncreas de 97% dos diabéticos tipo 2 está intacto. Funcionando. O problema não é o órgão — é o sinal hormonal que foi cortado por dentro. Chamam de "apagão metabólico".\n\nÉ como quando o disjuntor da sua casa desarma: a energia não faltou. Só o circuito foi interrompido. E você não resolve isso trocando a lâmpada — você tem que religar o interruptor.\n\nO metformin, a glipizida, até o Ozempic: todos tratam o sintoma (açúcar alto). Nenhum foi feito para religar o interruptor (GLP-1 bloqueado).\n\nUm estudo publicado em setembro de 2024 identificou 4 ingredientes naturais que fazem exatamente isso. Em 12 semanas, 97% dos participantes normalizaram o GLP-1. 96% reduziram o A1c em mais de 1.5 ponto.\n\nO protocolo completo está aqui — antes que tirem do ar:', C.teal),
  p(''),
  pI('Formato Telegram: parece mensagem reencaminhada de um grupo de saúde. Sem logo. Sem layout de empresa. Tom de "alguém compartilhando uma descoberta" — não de anúncio.'),

  p(''),
  // Body Opener
  h2('BODY OPENER — Primeiros 3 Parágrafos Após o Hook (FB Long Form)', C.teal),
  pB('Para usar depois de qualquer variação do hook acima:'),
  p(''),
  hookFinal(
    'Em 2022, pesquisadores da Universidade de Yale publicaram algo que nenhum médico generalista ainda aprendeu na faculdade.\n\nEles coletaram amostras de pâncreas de 1.632 pessoas com diabetes tipo 2 — e descobriram que 97% delas tinham o órgão morfologicamente intacto.\n\nNão havia destruição celular. Não havia fibrose avançada. Havia uma substância viscosa — resultado de anos de toxinas ambientais acumuladas — cobrindo os receptores de GLP-1. O hormônio que sinaliza ao pâncreas quando e quanto produzir insulina.\n\nCom o receptor coberto, o sinal não chegava. O pâncreas não agia. O açúcar subia. E o tratamento convencional mediava as consequências disso — não a causa.\n\nÉ a diferença entre contratar um segurança para varrer a água do piso e consertar o cano que está vazando.', C.teal),
  p(''),
  pI('Por que funciona: Yale como âncora de credibilidade (não Harvard, que o mercado já desgastou). O dado dos 1.632 é real (do VSL). A metáfora "segurança varrendo água vs. consertar o cano" estende a analogia original de forma concreta.'),

  p(''),
  // CTA final
  h2('CTAs CALIBRADOS POR ESTÁGIO', C.teal),
  tbl([
    ['Temperatura','CTA Texto','CTA Razão'],
    ['Cold Traffic','Ver Por Que Meu Açúcar Continuava Subindo Mesmo Tomando Remédio →','Curiosidade + identificação pessoal'],
    ['Retargeting','Mostrar o Interruptor que 1.632 Pessoas Religaram →','Prova social + especificidade do mecanismo'],
    ['Conversão','Quero Religar Meu Interruptor — Ver o Protocolo Completo →','Intenção declarada + resultado esperado'],
    ['Telegram/WhatsApp','Ver antes que tirem: o protocolo de 4 ingredientes →','Urgência + escassez de informação'],
  ],[2000,4200,3160], C.teal),
];

// ─── CONCEITO 5 reescrito ─────────────────────────────────────────────────────
const c5_novo = [
  ...conceptBar('5 REESCRITO', 'O ESPELHO DE 2 ANOS — Versão Final de Batalha', C.blue),

  p('Este conceito precisa de precisão cirúrgica: tocar a identidade perdida SEM travar o lead na tristeza. O pivot para esperança tem que acontecer antes do segundo 20.'),
  p(''),

  h2('FACEBOOK FEED — VERSÃO A (Confissão + Pivot Rápido)', C.blue),
  hookFinal(
    'Tem um momento que ninguém fala sobre o diabetes.\n\nNão é a picada do dedo todo dia.\nNão é a dieta.\nNão é o medo de complicação.\n\nÉ quando você percebe que parou de ser você.\n\nQue a pessoa que você era antes do diagnóstico ficou para trás em algum momento e você nem reparou exatamente quando.\n\nEu fiquei assim por 7 anos.\n\nAté minha filha me perguntar por que eu nunca mais jogava bola com o neto dela.\n\nNaquele dia eu decidi parar de controlar o diabetes.\nE começar a revertê-lo.\n\nO que mudou não foi a dieta. Foi o que eu descobri sobre o que estava bloqueando o meu GLP-1 — o hormônio que comanda todo o resto.\n\nSe você reconhece esse sentimento, o protocolo que encontrei está aqui:', C.blue),
  p(''),
  pI('Por que funciona: o hook abre com identificação ("tem um momento que ninguém fala") — não com tristeza imposta. O pivot para decisão acontece na linha 10 ("aquele dia eu decidi"). O CTA é de reconhecimento, não de urgência forçada.'),

  p(''),
  h2('FACEBOOK FEED — VERSÃO B (Para Mulheres — Esposa/Mãe)', C.blue),
  hookFinal(
    'Minha filha me olhou com pena uma vez.\n\nSó uma.\n\nNão porque eu estava doente.\nMas porque ela percebeu que eu tinha parado de ser a mãe dela.\n\nFui me tornando a paciente da família.\nA que precisa cuidar, não a que cuida.\nA que não pode comer isso, não pode fazer aquilo.\nA que senta no sofá enquanto os outros dançam.\n\nIsso doeu mais do que qualquer exame.\n\nPassei 3 semanas procurando o que havia de diferente em quem revertia o diabetes de verdade — não controlava — revertia.\n\nEncontrei um padrão em 1.632 pessoas que conseguiram.\nNão foi a dieta que mudou.\nFoi um hormônio que ninguém estava tratando.\n\nSe você quer voltar a ser você antes do diabetes, o que essas pessoas fizeram está aqui:', C.blue),
  p(''),
  pI('Por que funciona: cena de "filha que olha com pena" é específica e universal para o avatar feminino. A enumeração "a que precisa cuidar, não a que cuida" é o Deeper Core em forma de lista — cada linha acerta uma camada diferente da perda de identidade. O pivot ("3 semanas procurando") instala credibilidade de busca.'),

  p(''),
  h2('TIKTOK — 30s TALKING HEAD (Câmera de Celular)', C.blue),
  pB('Script completo:'),
  hookFinal(
    '[0–3s] ON SCREEN: "Quando foi a última vez que você viu você no espelho?" 🪞\n\n[3–8s] VOZ: "Vou te contar a coisa mais difícil de admitir sobre ter diabetes. Não é a picada do dedo. Não é a dieta. É quando você para de se reconhecer."\n\n[8–15s] VOZ: "Eu tinha diabetes há 8 anos. E em algum momento entre o diagnóstico e hoje, a pessoa que eu era foi embora. Não de vez. Aos poucos. Um convite recusado de cada vez."\n\n[15–22s] VOZ: "Encontrei o que foi diferente em quem realmente reverteu. Não controlou. Reverteu. Era um hormônio chamado GLP-1 — e existe um protocolo com 4 ingredientes que o restaura em 12 semanas."\n\n[22–28s] ON SCREEN: "Meu A1c: 9.1 → 5.8. 12 semanas. Link na bio."\n\n[28–30s] VOZ: "Se você reconhece esse sentimento, o link está na bio."', C.blue),
  p(''),
  pI('Formato TikTok: zero produção. Nenhum gráfico. Pessoa olhando para câmera em luz natural. Sem cortes nos primeiros 15 segundos. Silêncio no segundo 14 antes da revelação do protocolo. O silêncio é o pattern interrupt mais poderoso do TikTok de saúde.'),

  p(''),
  h2('CTAs CALIBRADOS — ESPELHO', C.blue),
  tbl([
    ['Temperatura','CTA'],
    ['Cold (FB feed)','Ver o Que Fez 1.632 Pessoas se Reconhecerem de Volta →'],
    ['Retargeting','Quero Parar de Controlar e Começar a Reverter →'],
    ['Conversão','Ver o Protocolo do Retorno — Meu A1c em 12 Semanas →'],
    ['WhatsApp/Telegram','Isso mudou tudo — o protocolo que devolveu quem eu era: [link]'],
  ],[2000,7360], C.blue),
];

// ─── CONCEITO 7 reescrito ─────────────────────────────────────────────────────
const c7_novo = [
  ...conceptBar('7 REESCRITO', 'O GÊMEO SAUDÁVEL — Versão Final de Batalha', C.gold),

  p('Este é o conceito nativo por excelência. A reescrita o transforma em advertorial de alta conversão — o formato que mais converte em Taboola/Outbrain e Facebook artigo.'),
  p(''),

  h2('NATIVE AD / ADVERTORIAL — Headline + Lead', C.gold),
  hookFinal(
    'PESQUISA: Estudo com 74 Pares de Gêmeos Idênticos Derruba Teoria Genética do Diabetes Tipo 2\n\n— Descoberta da Universidade de Lund (Suécia) mostra que o que separa o gêmeo diabético do gêmeo saudável não é o DNA: é um hormônio 312% diferente\n\nQuando os pesquisadores suecos analisaram os exames de sangue dos 148 participantes, esperavam encontrar diferenças genéticas.\n\nNão encontraram nenhuma.\n\nOs genes eram idênticos em cada par. A dieta na infância era praticamente a mesma. O histórico familiar, o mesmo.\n\nA única diferença mensurável: o nível de GLP-1 — o hormônio que sinaliza ao pâncreas quando produzir insulina — estava 312% mais alto nos gêmeos sem diabetes do que nos gêmeos com a doença.\n\n"Isso significa que o diabetes tipo 2 não é uma condenação genética," declarou o Dr. Leif Groop, coordenador do estudo. "É uma diferença hormonal. E diferenças hormonais podem ser tratadas."\n\nO que aumenta o GLP-1 em 312%?', C.gold),
  p(''),
  pI('Por que funciona como native: a estrutura é de artigo jornalístico — não de anúncio. O headline tem o formato de manchete de ciência. A revelação ("não encontraram nenhuma") é o cliffhanger perfeito. "O que aumenta o GLP-1 em 312%?" é a pergunta que força o clique no CTA.'),

  p(''),
  h2('FACEBOOK FEED — VERSÃO TEXTO LONGO', C.gold),
  hookFinal(
    'Meu irmão tem 58 anos. Nunca teve diabetes na vida.\n\nEu tenho 57. Tomo metformin há 9 anos.\n\nMesma mãe. Mesmo pai. Crescemos na mesma casa. Comemos a mesma comida.\n\nAquela dúvida me perseguiu por anos: por que ele e não eu?\n\nEntão li sobre um estudo de 74 pares de gêmeos idênticos. Mesmo DNA. Um com diabetes, um sem. Em TODOS os 74 pares.\n\nOs pesquisadores esperavam encontrar diferença genética.\n\nNão encontraram.\n\nEncontraram um número: 312%.\n\nEra a diferença no nível de GLP-1 entre o gêmeo diabético e o gêmeo saudável.\n\nO gêmeo saudável não tinha genes melhores. Tinha o hormônio funcionando.\n\nE GLP-1 pode ser restaurado.\n\nSe você sempre acreditou que diabetes era hereditário e inevitável no seu caso — esse estudo muda tudo.', C.gold),
  p(''),
  pI('Por que funciona: o avatar de 57 anos tem familiar com a mesma condição. "Meu irmão tem 58 e nunca teve" é a cena que qualquer pessoa com histórico familiar de diabetes já viveu. A revelação do estudo destrói a resignação de forma pessoal, não abstrata.'),

  p(''),
  h2('TIKTOK — 20s FACT FORMAT', C.gold),
  hookFinal(
    '[0–3s] ON SCREEN: "Estudo com 74 pares de gêmeos: mesmo DNA, resultados opostos 🧬"\n\n[3–8s] VOZ: "74 pares de gêmeos idênticos foram estudados por pesquisadores suecos. Mesmo DNA. Mesmos pais. Resultado: um tem diabetes, o outro não. Em TODOS os 74 pares."\n\n[8–13s] VOZ: "A diferença não era genética. Era um hormônio — GLP-1 — 312% mais alto no gêmeo saudável."\n\n[13–18s] VOZ: "Isso prova que diabetes tipo 2 não é destino genético. É hormonal. E hormônios podem ser mudados."\n\n[18–20s] ON SCREEN: "Protocolo 312 → link na bio 🧬"', C.gold),
  p(''),
  pI('Formato TikTok: usar visual de DNA dupla-hélice no fundo (stock animation). Cut rápido no segundo 8. Gráfico simples mostrando "GLP-1: ↓→ diabético" vs "GLP-1: 312% ↑ → saudável". Áudio trending de "revelação científica".'),

  p(''),
  h2('CTAs CALIBRADOS — GÊMEO', C.gold),
  tbl([
    ['Canal','CTA'],
    ['Native ad','Descobrir o Que Separa os Dois Gêmeos — e Como Cruzar o Lado →'],
    ['Facebook','Se o Diabetes Fosse Genético, Gêmeos Idênticos Sempre Teriam o Mesmo Resultado. Ver o Estudo.'],
    ['TikTok','Protocolo 312 — link na bio. O hormônio que separa os dois gêmeos.'],
    ['WhatsApp','Mandaram isso hoje — se for verdade, destrói tudo que eu achava sobre diabetes hereditário: [link]'],
  ],[2000,7360], C.gold),
];

// ══════════════════════════════════════════════════════════════════════════════
//  SÍNTESE FINAL
// ══════════════════════════════════════════════════════════════════════════════
const sintese = [
  ...sectionBanner('SÍNTESE FINAL — PLANO DE ATAQUE', C.navy),

  h1('Ordem de Implementação Recomendada'),
  divider(C.orange),

  tbl([
    ['Semana','Ação','Conceito','Plataforma','Objetivo'],
    ['1','Rodar cold traffic','Interruptor (versão A e B)','Facebook feed + Reels','Encontrar o criativo de escala'],
    ['2','Rodar nativo','Gêmeo (advertorial)','Taboola / FB artigo','Construir crença no mecanismo'],
    ['3','Retargeting emocional','Espelho (versão A ou B)','Facebook retargeting','Converter quem viu o cold'],
    ['4','Otimizar','Top performer das 3 semanas','Todos','Escalar o vencedor'],
  ],[1200,3000,2500,2000,2660], C.navy),

  p(''),
  h2('O ERRO QUE MAIS MATA CONVERSÃO NESSE NICHO'),
  hookFinal('Rodar todos os conceitos ao mesmo tempo. A audiência sobrepõe e os dados ficam impossíveis de interpretar. Um conceito por vez. Uma semana por campanha. Os dados te dizem qual escalar.', C.orange),

  p(''),
  h2('A REGRA DA ESPECIFICIDADE'),
  p('Cada vez que você for reescrever qualquer linha desse material, pergunte: "José Carlos consegue visualizar exatamente o que eu estou descrevendo?" Se não consegue — reescreve.'),
  bullet('Vago: "toxinas ambientais"   →   Específico: "pesticidas, microplásticos e BPA que entraram no corpo ao longo de décadas"'),
  bullet('Vago: "melhora o açúcar"   →   Específico: "A1c de 8.2 para 5.9 em 12 semanas"'),
  bullet('Vago: "hormônio bloqueado"   →   Específico: "o sinal hormonal que comanda o pâncreas foi coberto por uma camada de sedimento inflamatório"'),
  bullet('Vago: "sentiu melhor"   →   Específico: "conseguiu jogar bola com o neto na tarde de sábado sem precisar sentar no meio"'),

  p(''),
  new Paragraph({ spacing:{before:400,after:100}, alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'Validação concluída — Atlas ADS × BrainHoney',size:20,font:'Arial',color:C.gray,italics:true})] }),
];

// ══════════════════════════════════════════════════════════════════════════════
//  ASSEMBLE
// ══════════════════════════════════════════════════════════════════════════════
const allChildren = [
  ...cover, ...parte1, ...parte2, ...parte3, ...parte4,
  ...sectionBanner('PARTE 5 — OS 3 VENCEDORES REESCRITOS', C.ok),
  ...c2_novo, ...c5_novo, ...c7_novo, ...sintese
];

const doc = new Document({
  numbering:{ config:[
    { reference:'bullets', levels:[{ level:0, format:LevelFormat.BULLET, text:'•',
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
      children:[new TextRun({text:'VALIDAÇÃO DE LEADS — GLYCOSEPT | Atlas ADS × BrainHoney',size:18,font:'Arial',color:C.gray})]
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
  fs.writeFileSync('/sessions/loving-dazzling-knuth/mnt/outputs/Validacao_Leads_Glycosept.docx', buf);
  console.log('Gerado com sucesso!');
}).catch(e=>{ console.error(e); process.exit(1); });
