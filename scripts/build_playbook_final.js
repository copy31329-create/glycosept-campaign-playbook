const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, Header, Footer, PageBreak
} = require('docx');
const fs = require('fs');

// ─── PALETTE ─────────────────────────────────────────────────────────────────
const C = {
  navy:'1F3864', blue:'2E75B6', orange:'C55A11', red:'BF0000',
  green:'215732', gray:'595959', white:'FFFFFF', gold:'9A6F00',
  teal:'155A5A', purple:'4A235A', dark:'111111', mid:'707070',
  // concept colors
  C2:'155A5A',   // Interruptor — teal escuro
  C5:'1A3A6B',   // Espelho    — navy azul
  C7:'7A5200',   // Gêmeo      — âmbar escuro
  // audience expansion ring colors
  R0:'BF0000', R1:'C55A11', R2:'9A6F00', R3:'215732', R4:'1A3A6B', R5:'4A235A',
};
const CW=9360;
const CELL_M={top:100,bottom:100,left:160,right:160};
const BL={style:BorderStyle.SINGLE,size:1,color:'CCCCCC'};
const BORDERS={top:BL,bottom:BL,left:BL,right:BL};
const NB={top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const pb  = () => new Paragraph({children:[new PageBreak()]});
const sp  = (n=1) => new Paragraph({spacing:{before:n*80,after:n*80},children:[new TextRun({text:'',size:22,font:'Arial'})]});

function p(text,extra={}){return new Paragraph({spacing:{before:60,after:80},children:[new TextRun({text,size:22,font:'Arial',...extra})]})}
function pB(text,color=C.dark){return p(text,{bold:true,color})}
function pI(text,color=C.gray){return p(text,{italics:true,color,size:20})}
function pC(text,extra={}){return new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60,after:80},children:[new TextRun({text,size:22,font:'Arial',...extra})]})}

function divider(color=C.navy,thick=6){
  return new Paragraph({spacing:{before:180,after:180},border:{bottom:{style:BorderStyle.SINGLE,size:thick,color,space:1}},children:[]});
}

function bullet(text,bold_prefix=''){
  const ch=bold_prefix
    ?[new TextRun({text:bold_prefix+' ',bold:true,size:22,font:'Arial',color:C.dark}),new TextRun({text,size:22,font:'Arial'})]
    :[new TextRun({text,size:22,font:'Arial'})];
  return new Paragraph({spacing:{before:40,after:40},numbering:{reference:'bullets',level:0},children:ch});
}

// BIG NUMBER — estatística isolada, fonte enorme
function bigStat(number, label, color=C.navy){
  return [
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:120,after:0},
      children:[new TextRun({text:number,bold:true,size:96,font:'Arial',color})]}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:200},
      children:[new TextRun({text:label,size:22,font:'Arial',color:C.gray,italics:true})]}),
  ];
}

// FULLWIDTH COLOR BLOCK — como seção de brand guide
function colorBlock(text, fill, textColor=C.white, size=28){
  return new Paragraph({
    spacing:{before:0,after:0},
    shading:{fill,type:ShadingType.CLEAR},
    children:[new TextRun({text:`  ${text}  `,bold:true,size,font:'Arial',color:textColor})]
  });
}
function colorBlockSub(text, fill, textColor='DDDDDD', size=22){
  return new Paragraph({
    spacing:{before:0,after:280},
    shading:{fill,type:ShadingType.CLEAR},
    children:[new TextRun({text:`  ${text}  `,size,font:'Arial',color:textColor,italics:true})]
  });
}

// HOOK BLOCK
function hookBlock(text, accent){
  return new Paragraph({
    spacing:{before:120,after:120}, indent:{left:600},
    border:{left:{style:BorderStyle.SINGLE,size:28,color:accent,space:10}},
    shading:{fill:'F8F8F8',type:ShadingType.CLEAR},
    children:[new TextRun({text,size:22,font:'Arial',color:C.dark})]
  });
}

// DO / DON'T inline
function doRow(label,ok,bad,colWidths){
  return new Table({width:{size:CW,type:WidthType.DXA},columnWidths:colWidths,
    rows:[new TableRow({children:[
      new TableCell({borders:BORDERS,width:{size:colWidths[0],type:WidthType.DXA},margins:CELL_M,
        shading:{fill:'F0F0F0',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
        children:[new Paragraph({children:[new TextRun({text:label,bold:true,size:20,font:'Arial',color:C.gray})]})]}),
      new TableCell({borders:BORDERS,width:{size:colWidths[1],type:WidthType.DXA},margins:CELL_M,
        shading:{fill:'FFF0F0',type:ShadingType.CLEAR},
        children:[new Paragraph({children:[new TextRun({text:'✗  '+bad,size:20,font:'Arial',color:C.red})]})]}),
      new TableCell({borders:BORDERS,width:{size:colWidths[2],type:WidthType.DXA},margins:CELL_M,
        shading:{fill:'F0FFF4',type:ShadingType.CLEAR},
        children:[new Paragraph({children:[new TextRun({text:'✓  '+ok,size:20,font:'Arial',color:C.green})]})]}),
    ]})]
  });
}

function tbl(rows,colWidths,hFill=C.navy){
  return new Table({width:{size:CW,type:WidthType.DXA},columnWidths:colWidths,
    rows:rows.map((row,ri)=>new TableRow({tableHeader:ri===0,
      children:row.map((cell,ci)=>new TableCell({borders:BORDERS,
        width:{size:colWidths[ci],type:WidthType.DXA},margins:CELL_M,
        shading:ri===0?{fill:hFill,type:ShadingType.CLEAR}:(ri%2===0?{fill:'F8F8F8',type:ShadingType.CLEAR}:{fill:C.white,type:ShadingType.CLEAR}),
        verticalAlign:VerticalAlign.TOP,
        children:Array.isArray(cell)?cell:[new Paragraph({spacing:{before:40,after:40},
          children:[new TextRun({text:String(cell),size:20,font:'Arial',bold:ri===0,color:ri===0?C.white:C.dark})]})]}))})
    )});
}

// AUDIENCE CARD — card visual de público
function audienceCard(emoji, ring, title, tagline, hook, colWidth, fill, accent){
  return new TableCell({
    borders:BORDERS,
    width:{size:colWidth,type:WidthType.DXA},
    margins:{top:160,bottom:160,left:160,right:160},
    shading:{fill,type:ShadingType.CLEAR},
    verticalAlign:VerticalAlign.TOP,
    children:[
      new Paragraph({spacing:{before:0,after:40},children:[new TextRun({text:emoji,size:36,font:'Arial'})]}),
      new Paragraph({spacing:{before:0,after:20},children:[new TextRun({text:ring,size:18,font:'Arial',color:accent,bold:true})]}),
      new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:title,size:24,font:'Arial',bold:true,color:C.dark})]}),
      new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:tagline,size:19,font:'Arial',color:C.gray,italics:true})]}),
      new Paragraph({spacing:{before:60,after:0},children:[new TextRun({text:'Ângulo: ',bold:true,size:19,font:'Arial',color:accent})]}),
      new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:hook,size:19,font:'Arial',color:C.dark})]}),
    ]
  });
}

// CONCEPT PAGE BLOCK — 1 seção por conceito, estilo magazine
function conceptSection(num, name, tagline, accentColor,
  nomeName, nomeJustificativa,
  hookFB, hookTT, hookTG,
  emotion, cta, platform){
  return [
    pb(),
    // — Header bar —
    colorBlock(`CONCEITO ${num} — ${name}`, accentColor, C.white, 32),
    colorBlockSub(tagline, accentColor, 'DDDDDD', 20),

    // — Nome Chiclete —
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[2600,6760],
      rows:[new TableRow({children:[
        new TableCell({borders:BORDERS,width:{size:2600,type:WidthType.DXA},margins:CELL_M,
          shading:{fill:accentColor,type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
          children:[
            new Paragraph({spacing:{before:20,after:0},children:[new TextRun({text:'NOME CHICLETE',size:17,font:'Arial',color:'AAAAAA',bold:true})]}),
            new Paragraph({spacing:{before:0,after:20},children:[new TextRun({text:nomeName,size:30,font:'Arial',bold:true,color:C.white})]}),
          ]}),
        new TableCell({borders:BORDERS,width:{size:6760,type:WidthType.DXA},margins:CELL_M,
          shading:{fill:'F5F5F5',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
          children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:nomeJustificativa,size:20,font:'Arial',color:C.gray,italics:true})]})]})
      ]})]
    }),

    sp(),

    // — Hook por plataforma —
    pB('HOOKS POR PLATAFORMA', accentColor),

    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[1400,7960],
      rows:[
        // FB
        new TableRow({children:[
          new TableCell({borders:BORDERS,width:{size:1400,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'1877F2',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
            children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'f  Facebook',size:20,font:'Arial',bold:true,color:C.white})]})]}),
          new TableCell({borders:BORDERS,width:{size:7960,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'F0F8FF',type:ShadingType.CLEAR},
            children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:hookFB,size:21,font:'Arial',color:C.dark,italics:true})]})]})
        ]}),
        // TT
        new TableRow({children:[
          new TableCell({borders:BORDERS,width:{size:1400,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'000000',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
            children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'TT  TikTok',size:20,font:'Arial',bold:true,color:C.white})]})]}),
          new TableCell({borders:BORDERS,width:{size:7960,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'F5F5F5',type:ShadingType.CLEAR},
            children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:hookTT,size:21,font:'Arial',color:C.dark,italics:true})]})]})
        ]}),
        // TG
        new TableRow({children:[
          new TableCell({borders:BORDERS,width:{size:1400,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'2CA5E0',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
            children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'TG  Telegram',size:20,font:'Arial',bold:true,color:C.white})]})]}),
          new TableCell({borders:BORDERS,width:{size:7960,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'F0FAFF',type:ShadingType.CLEAR},
            children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:hookTG,size:21,font:'Arial',color:C.dark,italics:true})]})]})
        ]}),
      ]
    }),

    sp(),

    // — Emoção + CTA + Plataforma —
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[3120,3120,3120],
      rows:[
        new TableRow({children:[
          new TableCell({borders:BORDERS,width:{size:3120,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:accentColor,type:ShadingType.CLEAR},
            children:[new Paragraph({spacing:{before:20,after:0},children:[new TextRun({text:'EMOÇÃO PRINCIPAL',size:17,font:'Arial',bold:true,color:'AAAAAA'})]}),
              new Paragraph({spacing:{before:0,after:20},children:[new TextRun({text:emotion,size:21,font:'Arial',bold:true,color:C.white})]})]}),
          new TableCell({borders:BORDERS,width:{size:3120,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'F5F5F5',type:ShadingType.CLEAR},
            children:[new Paragraph({spacing:{before:20,after:0},children:[new TextRun({text:'CTA VENCEDOR',size:17,font:'Arial',bold:true,color:C.gray})]}),
              new Paragraph({spacing:{before:0,after:20},children:[new TextRun({text:cta,size:21,font:'Arial',bold:true,color:C.dark})]})]}),
          new TableCell({borders:BORDERS,width:{size:3120,type:WidthType.DXA},margins:CELL_M,
            shading:{fill:'F5F5F5',type:ShadingType.CLEAR},
            children:[new Paragraph({spacing:{before:20,after:0},children:[new TextRun({text:'PLATAFORMA IDEAL',size:17,font:'Arial',bold:true,color:C.gray})]}),
              new Paragraph({spacing:{before:0,after:20},children:[new TextRun({text:platform,size:21,font:'Arial',bold:true,color:C.dark})]})]}),
        ]})]
    }),
  ];
}

// ══════════════════════════════════════════════════════════════════════════════
//  1. CAPA
// ══════════════════════════════════════════════════════════════════════════════
const cover = [
  new Paragraph({spacing:{before:1800,after:0},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'GLYCOSEPT',bold:true,size:100,font:'Arial',color:C.navy})]}),
  new Paragraph({spacing:{before:0,after:0},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'CAMPAIGN PLAYBOOK',bold:true,size:48,font:'Arial',color:C.orange})]}),
  divider(C.orange,12),
  new Paragraph({spacing:{before:160,after:80},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:'3 conceitos validados · 6 públicos mapeados · hooks prontos para veicular',size:24,font:'Arial',color:C.gray,italics:true})]}),
  new Paragraph({spacing:{before:60,after:400},alignment:AlignmentType.CENTER,
    children:[new TextRun({text:`Atlas ADS × BrainHoney  |  ${new Date().toLocaleDateString('pt-BR')}`,size:20,font:'Arial',color:C.mid})]}),

  // 3 stats centrais
  new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[3120,3120,3120],
    rows:[new TableRow({children:[
      new TableCell({borders:NB,width:{size:3120,type:WidthType.DXA},margins:{top:200,bottom:200,left:200,right:200},
        shading:{fill:'F0F8FF',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'1.632',bold:true,size:80,font:'Arial',color:C.C2})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'participantes no estudo',size:19,font:'Arial',color:C.gray,italics:true})]}),
        ]}),
      new TableCell({borders:NB,width:{size:3120,type:WidthType.DXA},margins:{top:200,bottom:200,left:200,right:200},
        shading:{fill:'FFF8F0',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'97%',bold:true,size:80,font:'Arial',color:C.orange})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'normalizaram o GLP-1',size:19,font:'Arial',color:C.gray,italics:true})]}),
        ]}),
      new TableCell({borders:NB,width:{size:3120,type:WidthType.DXA},margins:{top:200,bottom:200,left:200,right:200},
        shading:{fill:'F0FFF4',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'312%',bold:true,size:80,font:'Arial',color:C.C7})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'diferença de GLP-1 no estudo dos gêmeos',size:19,font:'Arial',color:C.gray,italics:true})]}),
        ]}),
    ]})]
  }),

  sp(2),
  // Índice visual
  tbl([
    ['#','SEÇÃO','O QUE VOCÊ VAI ENCONTRAR'],
    ['01','O Produto','O mecanismo, a prova, o preço — em 30 segundos'],
    ['02','O Avatar','Quem é José Carlos — o lead que recebe os anúncios'],
    ['03','Mapa de Públicos','6 públicos mapeados do núcleo à expansão'],
    ['04','Conceito 2','O Interruptor Apagado — Protocolo GLP-Reset'],
    ['05','Conceito 5','O Espelho de 2 Anos — Método 1632'],
    ['06','Conceito 7','O Gêmeo Saudável — Protocolo 312'],
    ['07','Funil de Tráfego','Quem vê o quê, em qual ordem'],
    ['08','Implementação','O que rodar na semana 1, 2, 3 e 4'],
    ['09','Checklist','5 perguntas antes de qualquer criativo ir ao ar'],
  ],[600,2400,6360],C.navy),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  2. O PRODUTO EM 30 SEGUNDOS
// ══════════════════════════════════════════════════════════════════════════════
const produto = [
  colorBlock('01 — O PRODUTO', C.navy),
  colorBlockSub('Glycosept — tudo que você precisa saber em 30 segundos', C.navy),

  tbl([
    ['O QUÊ','DETALHE'],
    ['Produto','Glycosept — cápsula diária de 4 ingredientes naturais'],
    ['Ingredientes','MGO Manuka Honey · Quercetina · Pinosembrina · Geleia Real'],
    ['Mecanismo (MUP)','Lama tóxica (sedimento inflamatório) que cobre os receptores de GLP-1 no pâncreas'],
    ['Mecanismo (MUS)','Os 4 ingredientes dissolvem a lama e restauram o sinal GLP-1 — o pâncreas volta a responder'],
    ['Prova clínica','Estudo duplo-cego · 1.632 participantes · Journal of Metabolic Health'],
    ['Resultado #1','97% normalizaram GLP-1 · GLP-1 subiu +312%'],
    ['Resultado #2','96% reduziram A1c em mais de 1.5 ponto em 12 semanas'],
    ['Resultado #3','83% eliminaram ou reduziram medicação oral'],
    ['Garantia','60 dias — devolução total sem perguntas'],
    ['Preço','R$49/frasco (kit 6 frascos)'],
    ['Autoridade','Dr. William Lee (Harvard) · Dr. Gerald Shulman (Yale, 35 anos)'],
  ],[3000,6360],C.navy),

  sp(),
  // NÃO FAÇA / FAÇA
  pB('COMO POSICIONAR O PRODUTO', C.navy),
  sp(0.5),
  doRow('Posicionamento',
    'Suplemento que controla o açúcar',
    'Suplemento que resolve o sintoma',
    [2200,3580,3580]),
  sp(0.5),
  doRow('Promessa',
    'Reverte o bloqueio do GLP-1 — o pâncreas volta a funcionar',
    'Baixa o açúcar no sangue',
    [2200,3580,3580]),
  sp(0.5),
  doRow('Diferencial',
    'Age na causa (sinal GLP-1) — não no sintoma (glicemia alta)',
    'Natural e sem efeitos colaterais',
    [2200,3580,3580]),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  3. O AVATAR
// ══════════════════════════════════════════════════════════════════════════════
const avatar = [
  colorBlock('02 — O AVATAR', C.C2),
  colorBlockSub('José Carlos — o lead que vai receber cada anúncio', C.C2),

  // Stats do avatar em linha
  new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[1560,1560,1560,1560,1560,1560],
    rows:[new TableRow({children:[
      ...[
        ['57','anos'],['8.2','A1c atual'],['9','anos de metformin'],
        ['R$280','gasto mensal','em medicação'],['2','netos'],['1','cunhado amputado'],
      ].map(([n,l,l2])=>new TableCell({borders:BORDERS,width:{size:1560,type:WidthType.DXA},
        margins:{top:120,bottom:120,left:80,right:80},
        shading:{fill:'F0F8FF',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:n,bold:true,size:52,font:'Arial',color:C.C2})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:l,size:18,font:'Arial',color:C.gray,italics:true})]}),
          ...(l2?[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:l2,size:18,font:'Arial',color:C.gray,italics:true})]})]:[]),
        ]}))
    ]})]
  }),

  sp(),
  new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[4680,4680],
    rows:[new TableRow({children:[
      new TableCell({borders:BORDERS,width:{size:4680,type:WidthType.DXA},margins:CELL_M,
        shading:{fill:'FFF0F0',type:ShadingType.CLEAR},
        children:[
          new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:'O QUE FAZ ELE SAIR DO ANÚNCIO',bold:true,size:20,font:'Arial',color:C.red})]}),
          ...[
            'Nome de produto em inglês (Grizzly, Protocol X)',
            'Promessa absurda nos primeiros 3 segundos',
            'Depoimento com nome americano (Margaret, John)',
            'Layout que parece empresa — logo, banner, jaleco',
            'Linguagem de conspiração explícita (Pfizer roubou...)',
            '"Cura o diabetes" — dispara detector de fraude',
          ].map(t=>new Paragraph({spacing:{before:30,after:30},numbering:{reference:'bullets',level:0},
            children:[new TextRun({text:t,size:20,font:'Arial',color:C.dark})]})),
        ]}),
      new TableCell({borders:BORDERS,width:{size:4680,type:WidthType.DXA},margins:CELL_M,
        shading:{fill:'F0FFF4',type:ShadingType.CLEAR},
        children:[
          new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:'O QUE FAZ ELE CLICAR',bold:true,size:20,font:'Arial',color:C.green})]}),
          ...[
            'Cena que ele já viveu (neto, esposa, exame na mesa)',
            'Inversão de culpa — não é fraqueza, é causa externa',
            'Analogia cotidiana (disjuntor, cano entupido)',
            'Número específico — 312%, 1.632, 97%, 8.4→5.9',
            'Depoimento com nome e cidade BR (Dona Amélia, SP)',
            'Voz de primeira pessoa — parece post, não anúncio',
          ].map(t=>new Paragraph({spacing:{before:30,after:30},numbering:{reference:'bullets',level:0},
            children:[new TextRun({text:t,size:20,font:'Arial',color:C.dark})]})),
        ]}),
    ]})]
  }),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  4. MAPA DE PÚBLICOS — 6 anéis
// ══════════════════════════════════════════════════════════════════════════════
const publicos = [
  colorBlock('03 — MAPA DE PÚBLICOS', C.navy),
  colorBlockSub('Do núcleo à expansão — 6 públicos, mesmo produto, ângulos diferentes', C.navy),

  sp(),
  pB('COMO LER ESTE MAPA', C.gray),
  p('Núcleo = público atual. Cada anel externo = mais leads disponíveis com pequeno ajuste de ângulo no hook. O produto e a oferta não mudam.'),
  sp(),

  // Tabela de anéis
  tbl([
    ['ANEL','PÚBLICO','TAMANHO','EMOÇÃO CENTRAL','ÂNGULO DE ENTRADA'],
    ['🔴 NÚCLEO','Diabético tipo 2 diagnosticado\n50+, metformin, A1c fora de controle',
     '16M no Brasil','Esperança + indignação','"Não é culpa sua. O sinal foi cortado."'],
    ['🟠 ANEL 1','Pré-diabético\nA1c 5.7–6.4, ainda sem remédio',
     '3–4x maior que o núcleo','Medo preventivo','"Seu médico disse \'na beira\'. O GLP-1 já está caindo há anos."'],
    ['🟡 ANEL 2','Família que compra para o outro\nEsposa, filho, irmão',
     'Alto — comprador ≠ usuário','Dor de quem assiste','"Você está vendo alguém que ama desaparecer aos poucos."'],
    ['🟢 ANEL 3','Usuário de Ozempic buscando saída\nPreocupado com custo, efeito, dependência',
     'Crescente em 2025–2026','Indignação com custo','"O original: sem agulha, sem conta mensal, sem dependência."'],
    ['🔵 ANEL 4','Desistente\nParou de medir, de tentar, de acreditar',
     'Invisível — mas enorme','Exaustão validada','"Para quem já tentou tudo e parou de tentar. Sem julgamento."'],
    ['🟣 ANEL 5','Prevenção — histórico familiar\n50+, pai e avô tiveram, sem diagnóstico ainda',
     'Médio-alto','Medo de destino','"Meu pai tinha. Meu avô tinha. Eu decidi mudar o trajeto."'],
  ],[800,3200,2000,1700,1660],C.navy),

  sp(),
  pB('PRIORIDADE DE ATIVAÇÃO', C.navy),
  sp(0.5),
  tbl([
    ['FASE','ANEL','QUANDO ATIVAR','CRIATIVO BASE'],
    ['Fase 1 (agora)','Núcleo + Anel 1','Semanas 1–4','Conceitos 2, 5, 7 com hooks atuais'],
    ['Fase 2','Anel 2 (família)','Após validar o criativo de escala','Conceito 5 adaptado — "você vê ela sumir"'],
    ['Fase 3','Anel 3 (Ozempic)','Quando tiver prova social sólida','Conceito 3 adaptado — sem conspiração, só comparativo'],
    ['Fase 4','Anéis 4 e 5','Após 60–90 dias de dados','Tom completamente diferente — começar do zero'],
  ],[1000,2000,2760,3600],C.navy),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  5–7. CONCEITOS (1 seção cada)
// ══════════════════════════════════════════════════════════════════════════════
const conceito2 = conceptSection(
  '2','O INTERRUPTOR APAGADO',
  'O pâncreas está intacto. O sinal foi cortado. Aqui está como religá-lo.',
  C.C2,
  'Protocolo GLP-Reset',
  '"Reset" é universal: todo mundo sabe o que é resetar algo. GLP âncora no mecanismo real e verificável.',
  // FB
  'Fiz algo que nenhum médico me pediu.\nPeguei todos os meus exames dos últimos 9 anos e coloquei lado a lado na mesa.\nMetformin. A1c 8.4.\nMetformin. A1c 8.1.\nMetformin. A1c 8.6.\n9 anos. O mesmo remédio. O mesmo resultado.\nEntão perguntei pro meu médico: se esse remédio está funcionando, por que o A1c continua igual?\nEle ficou em silêncio por 4 segundos.\n"Porque o metformin trata o açúcar que sobrou. Não o motivo pelo qual o açúcar sobe."\nAquele silêncio mudou tudo.',
  // TikTok
  '[0–2s] "9 anos de metformin. Resultado: zero." | [2–5s] "Ninguém me contou que meu pâncreas estava intacto o tempo todo." | [5–8s] "O problema nunca foi o órgão. Era o sinal. 🔌" | [CTA] Link na bio → Protocolo GLP-Reset',
  // Telegram
  '⚡ O pâncreas de 97% dos diabéticos tipo 2 está intacto. O problema é o sinal hormonal que foi cortado por dentro — chamam de apagão metabólico. Um estudo com 1.632 pessoas encontrou 4 ingredientes que religam esse sinal em 12 semanas. O protocolo completo: [link]',
  // emoção / cta / plataforma
  'Raiva justa\n→ Esperança técnica',
  'Ver por que meu A1c continuava subindo mesmo tomando remédio →',
  'Facebook Feed\nCold Traffic principal'
);

const conceito5 = conceptSection(
  '5','O ESPELHO DE 2 ANOS',
  'O diabetes não roubou só a saúde. Roubou quem você era. Isso pode ser revertido.',
  C.C5,
  'Método 1632',
  'O número exato do estudo virou o nome. Específico demais para ser inventado — credibilidade embutida.',
  // FB
  'Meu neto me perguntou por que eu nunca corria com ele.\nEu disse que estava cansado.\nEle virou as costas e foi brincar sozinho.\nEu fiquei olhando para ele correr, imóvel no meio do quintal.\nNão era dor física que me parava.\nEra a sensação de que eu tinha ficado velho antes da hora.\nNaquele dia eu parei de aceitar que "controlar" era o melhor que eu podia fazer.\nE encontrei o que realmente reverte — não controla — o que estava acontecendo.',
  // TikTok
  '[0–3s SILÊNCIO] "Quando você parou de ser você?" 🪞 | [3–8s] "Tem uma coisa sobre o diabetes que ninguém fala. Não é o A1c. É quando você para de se reconhecer." | [15–20s] "Decidi parar de controlar. E encontrei o que reverte." | [CTA] Método 1632 → link na bio',
  // Telegram
  '📱 Compartilhei isso com 3 pessoas que conheço com diabetes. A diferença não foi só nos exames — foi em quem elas voltaram a ser. 1.632 pessoas, 12 semanas, 96% reduziram o A1c mais de 1.5 ponto — sem mudar a dieta. O protocolo completo: [link]',
  'Identidade perdida\n→ Restauração',
  'Ver o que fez 1.632 pessoas se reconhecerem de volta →',
  'Facebook Feed + Retargeting\nEspecialmente mulheres'
);

const conceito7 = conceptSection(
  '7','O GÊMEO SAUDÁVEL',
  'Mesmo DNA. Um com diabetes, um sem. A diferença é um número: 312.',
  C.C7,
  'Protocolo 312',
  '2 palavras. O número real do estudo. Impossível de copiar sem citar a mesma fonte científica.',
  // FB
  'Meu irmão tem 61 anos. Nunca tomou remédio para diabetes na vida.\nCome arroz todo dia. A1c dele: 5.1.\nO meu, tomando metformin há 9 anos: 8.4.\nPassei anos achando que tinha herdado o "lado ruim" do sangue.\nEntão li sobre 74 pares de gêmeos idênticos — DNA literalmente igual.\nEm todos os 74: um com diabetes. Um sem.\nOs pesquisadores esperavam diferença genética.\nNão encontraram nenhuma.\nO que encontraram foi um número: 312.\nE hormônios podem ser mudados.',
  // TikTok
  '[0–2s] "74 pares de gêmeos: mesmo DNA, resultados opostos 🧬" | [2–6s] "Estudo sueco. Todos os 74 pares: um com diabetes, um sem. A diferença não era genética." | [6–10s] "Era um hormônio — 312% diferente entre os dois." | [CTA] Protocolo 312 → link na bio',
  // Telegram (native/artigo)
  '🧬 DESCOBERTA: Estudo com 74 pares de gêmeos idênticos derruba teoria genética do diabetes tipo 2. Pesquisadores esperavam diferença no DNA. Encontraram diferença de 312% em um hormônio. "Diabetes tipo 2 não é condenação genética — é hormonal. E hormônios mudam." O protocolo completo: [link]',
  'Injustiça\n→ Possibilidade',
  'Descobrir o que separa os dois gêmeos — e como cruzar o lado →',
  'Native Ads · Taboola\nFacebook Artigo'
);

// ══════════════════════════════════════════════════════════════════════════════
//  8. FUNIL DE TRÁFEGO
// ══════════════════════════════════════════════════════════════════════════════
const funil = [
  pb(),
  colorBlock('07 — FUNIL DE TRÁFEGO', C.navy),
  colorBlockSub('Quem vê o quê, em qual ordem — do primeiro contato à compra', C.navy),

  sp(),

  // Funil visual com 4 estágios
  new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[400,2200,2000,2560,2200],
    rows:[
      // Header
      new TableRow({tableHeader:true,children:[
        new TableCell({borders:BORDERS,width:{size:400,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'',size:18,font:'Arial',color:C.white})]}),]}),
        new TableCell({borders:BORDERS,width:{size:2200,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'ESTÁGIO',bold:true,size:20,font:'Arial',color:C.white})]})]}),
        new TableCell({borders:BORDERS,width:{size:2000,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'CONCEITO',bold:true,size:20,font:'Arial',color:C.white})]})]}),
        new TableCell({borders:BORDERS,width:{size:2560,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'PLATAFORMA',bold:true,size:20,font:'Arial',color:C.white})]})]}),
        new TableCell({borders:BORDERS,width:{size:2200,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'OBJETIVO',bold:true,size:20,font:'Arial',color:C.white})]})]}),
      ]}),
      // Linha 1 — TOPO
      ...[[
        '🥶','TOPO — Cold Traffic\n(nunca viu o produto)',
        'C2 — Interruptor\nC7 — Gêmeo Saudável',
        'Facebook Reels · TikTok\nNative Ads (Taboola)',
        'Parar o scroll.\nPlantar a dúvida.\nNão mencionar produto.'
      ],[
        '🌡','MEIO — Educação\n(viu, ficou curioso)',
        'C7 — Gêmeo (advertorial)\nC2 — Interruptor (longo)',
        'Facebook Feed · YouTube\nRetargeting',
        'Construir crença\nno mecanismo GLP-1.'
      ],[
        '🔥','FUNDO — Identidade\n(está considerando)',
        'C5 — Espelho\n(Método 1632)',
        'Facebook Retargeting\nTelegram · WhatsApp',
        'Tocar o Deeper Core.\nCrear urgência emocional.'
      ],[
        '💳','CONVERSÃO\n(pronto para comprar)',
        'VSL Completo (Glycosept)\nDepoimentos + Oferta',
        'Landing Page\n(tráfego de todos os canais)',
        'Fechar com prova\n+ urgência real (60 dias).'
      ]].map(([e,s,c,pl,o])=>new TableRow({children:[
        new TableCell({borders:BORDERS,width:{size:400,type:WidthType.DXA},margins:CELL_M,shading:{fill:'F8F8F8',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:e,size:32,font:'Arial'})]}),]}),
        new TableCell({borders:BORDERS,width:{size:2200,type:WidthType.DXA},margins:CELL_M,shading:{fill:'F8F8F8',type:ShadingType.CLEAR},children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:s,size:20,font:'Arial',color:C.dark})]})]}),
        new TableCell({borders:BORDERS,width:{size:2000,type:WidthType.DXA},margins:CELL_M,shading:{fill:'F0F8FF',type:ShadingType.CLEAR},children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:c,size:20,font:'Arial',color:C.C2})]})]}),
        new TableCell({borders:BORDERS,width:{size:2560,type:WidthType.DXA},margins:CELL_M,shading:{fill:'F8F8F8',type:ShadingType.CLEAR},children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:pl,size:20,font:'Arial',color:C.dark})]})]}),
        new TableCell({borders:BORDERS,width:{size:2200,type:WidthType.DXA},margins:CELL_M,shading:{fill:'F8F8F8',type:ShadingType.CLEAR},children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:o,size:20,font:'Arial',color:C.gray,italics:true})]})]}),
      ]})),
    ]
  }),
  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  9. PLANO DE IMPLEMENTAÇÃO — 4 SEMANAS
// ══════════════════════════════════════════════════════════════════════════════
const implementacao = [
  colorBlock('08 — PLANO DE IMPLEMENTAÇÃO', C.navy),
  colorBlockSub('O que rodar na semana 1, 2, 3 e 4 — sem tentar tudo ao mesmo tempo', C.navy),

  sp(),
  p('REGRA DE OURO: um conceito por semana. Um criativo por campanha. Os dados dizem qual escalar.', {bold:true,color:C.red}),
  sp(),

  tbl([
    ['SEMANA','CONCEITO','HOOK','PLATAFORMA','META DA SEMANA'],
    ['Semana 1','C2 — Interruptor\n(Protocolo GLP-Reset)','Confissão pessoal:\n"9 anos. Mesmo remédio. Mesmo resultado."',
     'Facebook Feed + Reels\nCold traffic','Encontrar CTR acima de 1%.\nIdentificar faixa etária que mais converte.'],
    ['Semana 2','C7 — Gêmeo Saudável\n(Protocolo 312)','Irmão com 61 anos:\n"A1c 5.1. O meu: 8.4."',
     'Native ads (Taboola)\nFacebook Artigo','Testar advertorial longo.\nMedir scroll depth e VSL start rate.'],
    ['Semana 3','C5 — Espelho\n(Método 1632)','Neto correndo sozinho:\n"Imóvel no meio do quintal."',
     'Facebook Retargeting\n(quem viu S1 e S2)','Converter leads que não compraram.\nMede: CPP e ROAS do retargeting.'],
    ['Semana 4','Escalar o vencedor','Top performer das 3 semanas anteriores',
     'Ampliar budget no canal vencedor','Dobrar verba no criativo de menor CPA.\nIniciar expansão para Anel 1 (pré-diabéticos).'],
  ],[1200,2600,2800,1800,2960],C.navy),

  pb(),
];

// ══════════════════════════════════════════════════════════════════════════════
//  10. CHECKLIST PRÉ-VEICULAÇÃO
// ══════════════════════════════════════════════════════════════════════════════
const checklist = [
  colorBlock('09 — CHECKLIST PRÉ-VEICULAÇÃO', C.navy),
  colorBlockSub('5 perguntas. Se uma resposta for NÃO — o criativo não vai ao ar.', C.navy),

  sp(),

  new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[600,6800,1960],
    rows:[
      new TableRow({tableHeader:true,children:[
        new TableCell({borders:BORDERS,width:{size:600,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'#',bold:true,size:20,font:'Arial',color:C.white})]})]}),
        new TableCell({borders:BORDERS,width:{size:6800,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'PERGUNTA',bold:true,size:20,font:'Arial',color:C.white})]})]}),
        new TableCell({borders:BORDERS,width:{size:1960,type:WidthType.DXA},margins:CELL_M,shading:{fill:C.navy,type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:'COMO VERIFICAR',bold:true,size:20,font:'Arial',color:C.white})]})]}),
      ]}),
      ...[
        ['1','A primeira linha cria tensão não resolvida?\n(O lead quer saber o que vem depois — não recebeu informação logo)','Cubra tudo exceto a linha 1. Ela sozinha gera curiosidade?'],
        ['2','José Carlos se vê nessa cena em 3 segundos?\n(O lead específico — 55+, diabético, BR — se reconhece imediatamente)','Mostre para 3 pessoas do avatar. 2 disseram "isso aconteceu comigo"?'],
        ['3','O produto não aparece nos primeiros 5 parágrafos?\n(Começa com cena/dor — nunca com produto, nome ou benefício)','Conte os parágrafos. Produto antes do 6º = reescreve.'],
        ['4','A voz é de pessoa, não de empresa?\n(Soa como post pessoal, não como anúncio veiculado)','Troque "nosso produto" por "eu". O texto faz sentido?'],
        ['5','O nome chiclete aparece DEPOIS do mecanismo ser explicado?\n(O nome não precisa de explicação — ele confirma o que o lead já entendeu)','Onde o nome aparece pela primeira vez? Deve ser após a virada do mecanismo.'],
      ].map(([n,q,v],i)=>new TableRow({children:[
        new TableCell({borders:BORDERS,width:{size:600,type:WidthType.DXA},margins:CELL_M,
          shading:{fill:i%2===0?'F8F8F8':'FFFFFF',type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
          children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:n,bold:true,size:28,font:'Arial',color:C.orange})]})]}),
        new TableCell({borders:BORDERS,width:{size:6800,type:WidthType.DXA},margins:CELL_M,
          shading:{fill:i%2===0?'F8F8F8':'FFFFFF',type:ShadingType.CLEAR},
          children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:q,size:21,font:'Arial',color:C.dark})]})]}),
        new TableCell({borders:BORDERS,width:{size:1960,type:WidthType.DXA},margins:CELL_M,
          shading:{fill:i%2===0?'F0FFF4':'F8FFF8',type:ShadingType.CLEAR},
          children:[new Paragraph({spacing:{before:20,after:20},children:[new TextRun({text:v,size:18,font:'Arial',color:C.gray,italics:true})]})]}),
      ]})),
    ]
  }),

  sp(2),
  divider(C.orange,8),
  sp(),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:100,after:60},
    children:[new TextRun({text:'GLYCOSEPT CAMPAIGN PLAYBOOK',bold:true,size:28,font:'Arial',color:C.navy})]}),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:100},
    children:[new TextRun({text:'Atlas ADS × BrainHoney  |  Documento Confidencial',size:20,font:'Arial',color:C.gray,italics:true})]}),
];

// ══════════════════════════════════════════════════════════════════════════════
//  ASSEMBLE
// ══════════════════════════════════════════════════════════════════════════════
const allChildren = [
  ...cover,
  ...produto,
  ...avatar,
  ...publicos,
  ...conceito2,
  ...conceito5,
  ...conceito7,
  ...funil,
  ...implementacao,
  ...checklist,
];

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
    properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.orange,space:6}},
      children:[new TextRun({text:'GLYCOSEPT CAMPAIGN PLAYBOOK | Atlas ADS × BrainHoney',size:17,font:'Arial',color:C.mid})]
    })]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({text:'Pág. ',size:17,font:'Arial',color:C.mid}),
        new TextRun({children:[PageNumber.CURRENT],size:17,font:'Arial',color:C.mid}),
        new TextRun({text:'  |  Confidencial — Atlas ADS',size:17,font:'Arial',color:C.mid})]
    })]})},
    children:allChildren,
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('/sessions/loving-dazzling-knuth/mnt/outputs/Glycosept_Campaign_Playbook_FINAL.docx',buf);
  console.log('Gerado com sucesso!');
}).catch(e=>{console.error(e);process.exit(1);});
