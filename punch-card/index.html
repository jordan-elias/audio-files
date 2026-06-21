const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer, ExternalHyperlink,
  PageBreak
} = require('docx');
const fs = require('fs');

// ── Colour palette ──────────────────────────────────────────────
const C = {
  black:    '1A1A1A',
  darkGrey: '3A3A3A',
  midGrey:  '777777',
  lightGrey:'E8E8E8',
  rule:     'CCCCCC',
  accent:   '2A2A2A',
  white:    'FFFFFF',
};

// ── Helpers ─────────────────────────────────────────────────────
const sp = (before=0, after=0) => ({ spacing: { before, after } });
const ind = (left=0, hanging=0) => ({ indent: { left, hanging } });
const border = (color=C.rule, size=4) => ({ style: BorderStyle.SINGLE, size, color });
const allBorders = () => {
  const b = border(); return { top:b, bottom:b, left:b, right:b };
};
const noBorders = () => {
  const nb = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top:nb, bottom:nb, left:nb, right:nb };
};

function para(text, opts={}) {
  const {
    heading, bold=false, italic=false, size=22, color=C.black,
    spBefore=0, spAfter=160, align=AlignmentType.LEFT,
    font='Arial', indLeft=0
  } = opts;
  return new Paragraph({
    heading,
    alignment: align,
    ...sp(spBefore, spAfter),
    ...(indLeft ? ind(indLeft) : {}),
    children: [new TextRun({ text, bold, italic, size, color, font })]
  });
}

function rule() {
  return new Paragraph({
    ...sp(80, 80),
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.rule } },
    children: [new TextRun('')]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    ...sp(400, 160),
    children: [new TextRun({ text, bold: true, size: 32, font: 'Arial', color: C.black })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    ...sp(320, 120),
    children: [new TextRun({ text, bold: true, size: 26, font: 'Arial', color: C.black })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    ...sp(240, 80),
    children: [new TextRun({ text, bold: true, size: 22, font: 'Arial', color: C.darkGrey })]
  });
}

function body(text, opts={}) {
  return para(text, { size: 22, spAfter: 160, ...opts });
}

function caption(text) {
  return para(text, { size: 18, color: C.midGrey, italic: true, spAfter: 80 });
}

function bullet(text, sub=false) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: sub ? 1 : 0 },
    ...sp(0, 100),
    children: [new TextRun({ text, size: 22, font: 'Arial', color: C.black })]
  });
}

function twoColRow(label, value, shaded=false) {
  const fill = shaded ? 'F2F2F2' : C.white;
  const cellOpts = (w) => ({
    borders: allBorders(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
  });
  return new TableRow({ children: [
    new TableCell({ ...cellOpts(2800), children: [
      new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: 'Arial', color: C.darkGrey })] })
    ]}),
    new TableCell({ ...cellOpts(6560), children: [
      new Paragraph({ children: [new TextRun({ text: value, size: 20, font: 'Arial', color: C.black })] })
    ]}),
  ]});
}

function mappingRow(feature, sound, shaded=false) {
  const fill = shaded ? 'F2F2F2' : C.white;
  const c = (w) => ({
    borders: allBorders(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
  });
  return new TableRow({ children: [
    new TableCell({ ...c(4680), children: [
      new Paragraph({ children: [new TextRun({ text: feature, size: 20, font: 'Arial', color: C.black })] })
    ]}),
    new TableCell({ ...c(4680), children: [
      new Paragraph({ children: [new TextRun({ text: sound, size: 20, font: 'Arial', color: C.black })] })
    ]}),
  ]});
}

function threeColRow(a, b, c_, shaded=false) {
  const fill = shaded ? 'F2F2F2' : C.white;
  const cell = (w, text, bold=false) => new TableCell({
    borders: allBorders(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [new Paragraph({ children: [new TextRun({ text, bold, size: 20, font: 'Arial', color: C.black })] })]
  });
  return new TableRow({ children: [cell(2800,a,bold=false), cell(3200,b), cell(3360,c_)] });
}

function tableHeader2(l, r) {
  const cell = (w, text) => new TableCell({
    borders: allBorders(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill: C.black, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: 'Arial', color: C.white })] })]
  });
  return new TableRow({ children: [cell(4680,l), cell(4680,r)] });
}

function tableHeader3(a, b, c_) {
  const cell = (w, text) => new TableCell({
    borders: allBorders(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill: C.black, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: 'Arial', color: C.white })] })]
  });
  return new TableRow({ children: [cell(2800,a), cell(3200,b), cell(3360,c_)] });
}

function tableHeader4(cols) {
  return new TableRow({
    children: cols.map(([text, w]) => new TableCell({
      borders: allBorders(),
      width: { size: w, type: WidthType.DXA },
      shading: { fill: C.black, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: 'Arial', color: C.white })] })]
    }))
  });
}

function fourColRow(a, b, c_, d, shaded=false) {
  const fill = shaded ? 'F2F2F2' : C.white;
  const widths = [2200, 2200, 2480, 2480];
  const vals = [a, b, c_, d];
  return new TableRow({
    children: vals.map((text, i) => new TableCell({
      borders: allBorders(),
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text, size: 20, font: 'Arial', color: C.black })] })]
    }))
  });
}

// ── Document ────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: C.black },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial', color: C.black },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 22, bold: true, font: 'Arial', color: C.darkGrey },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '\u2014', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 560, hanging: 280 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1120, hanging: 280 } } } },
      ]},
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.rule } },
          ...sp(0, 120),
          children: [new TextRun({ text: 'Textile Sound Workshop — Build Plan', size: 18, color: C.midGrey, font: 'Arial' })]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.rule } },
          ...sp(120, 0),
          children: [
            new TextRun({ text: 'Jordan Elias  \u00B7  jordanelias.de  \u00B7  Page ', size: 18, color: C.midGrey, font: 'Arial' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.midGrey, font: 'Arial' }),
          ]
        })
      ]})
    },
    children: [

      // ══════════════════════════════════════════════════════════
      // COVER
      // ══════════════════════════════════════════════════════════
      new Paragraph({ ...sp(720, 80), children: [
        new TextRun({ text: 'TEXTILE SOUND WORKSHOP', bold: true, size: 56, font: 'Arial', color: C.black })
      ]}),
      new Paragraph({ ...sp(0, 40), children: [
        new TextRun({ text: 'A build plan for two instruments at the intersection of knitting and experimental music', size: 26, font: 'Arial', color: C.midGrey, italic: true })
      ]}),
      rule(),
      new Paragraph({ ...sp(80, 80), children: [
        new TextRun({ text: 'Jordan Elias, MT-BC  \u00B7  Berlin  \u00B7  2026', size: 20, font: 'Arial', color: C.midGrey })
      ]}),
      new Paragraph({ children: [new PageBreak()] }),

      // ══════════════════════════════════════════════════════════
      // CONCEPT
      // ══════════════════════════════════════════════════════════
      h1('Concept and Context'),
      body('This document outlines two interconnected hardware builds for a workshop bringing together knitting and experimental music. The workshop is designed for people who already knit garments and want a genuinely novel creative process — one that produces both a textile object and a musical experience from the same source material.'),
      body('The central premise is simple: a knitting machine punch card is a binary grid. Each hole is either punched or not. So is a MIDI file. So is a piece of sheet music. So is a player piano roll. The punch card is already a score — it just needs a reader that plays it as sound rather than fabric.'),
      body('Both projects described here are honest to that premise. They do not translate the card into conventional melodies. They use it as a spatial control surface for ambient and experimental sound — drones, harmonics, formants, resonant strings — in a direct correspondence between the visual structure of the pattern and its sonic character. A dense cluster of holes sounds dense. A symmetrical pattern sounds balanced. An empty row is silence.'),
      body('The two projects exist on a spectrum of complexity and spectacle. Project 1 is a standalone instrument — a hand-cranked reader you can build in a weekend and bring to any space. Project 2 is an integration with an existing machine — more involved, more conceptually powerful, and potentially the centrepiece of an ongoing series.'),
      new Paragraph({ children: [new PageBreak()] }),

      // ══════════════════════════════════════════════════════════
      // PROJECT 1
      // ══════════════════════════════════════════════════════════
      h1('Project 1: The Punch Card Synth'),
      body('A hand-cranked reader that plays a knitting punch card as ambient sound. Participants turn a crank, the card advances through a slot, and sound emerges from speakers. Faster cranking = faster music. Slower cranking = time stretches. The person becomes the tempo.'),
      body('This is a player piano for knitting patterns. The mechanism is simple enough to build from off-the-shelf components. The experience is immediate and intuitive for anyone who has never encountered either electronics or experimental music.', { spAfter: 240 }),

      h2('The Experience'),
      body('Imagine a wooden or acrylic enclosure roughly the size of a shoebox, sitting on a table. A slot in one end receives a standard knitting punch card — 24 columns, any number of rows. Inside, a row of 24 LEDs shines through the card onto 24 phototransistors. A hand crank on the side turns rubber rollers that advance the card one row at a time.'),
      body('As the card moves, the sensor array reads each row. Open holes (punched) allow light through. Closed positions block it. The Arduino reads this as a 24-bit binary number 30 times per second and sends the data to a laptop running SuperCollider, which maps the spatial distribution of holes to a continuously evolving ambient soundscape.'),
      body('The result is not a melody. It is a sonic landscape shaped by the pattern — its density, its symmetry, its clusters and voids. Drawing a diamond produces different sound from a chevron. A dense row of holes produces a thick harmonic cluster. An empty row drops to near-silence. People can hear the structure of their own pattern.', { spAfter: 240 }),

      h2('Hardware Components'),

      h3('The Reader'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Sensor array', '24 × LTR-301 phototransistors (or photoresistors) in a row, spaced 5 mm apart — matching standard Brother/Silver Reed punch card hole spacing'),
          twoColRow('Light source', '24 × 3mm warm-white LEDs in a parallel row on the opposite side of the card slot', true),
          twoColRow('Card slot', 'Machined acrylic or laser-cut plywood channel, 2 mm clearance each side, open top for loading'),
          twoColRow('Enclosure', 'Laser-cut 6 mm plywood or clear 5 mm acrylic — translucent acrylic allows the LED glow to be visible, adding to the aesthetic', true),
          twoColRow('Dimensions', 'Approximately 200 mm × 80 mm × 60 mm — small enough to sit on a knitting machine or table'),
        ]
      }),
      body(''),

      h3('The Transport Mechanism'),
      body('Standard punch cards are 24 columns wide and vary in length — typically 24–60 rows. The transport mechanism must advance the card one row at a time (each row is 5 mm, matching the hole spacing) and signal the microcontroller at each advance.'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Crank', 'Simple wooden or metal handle attached to a shaft — no gearing needed at first'),
          twoColRow('Rollers', '2 × rubber-coated rollers (O-ring on aluminium rod works well), one driven, one idler — pinch the card between them', true),
          twoColRow('Encoder', 'Bourns PEC11R rotary encoder on the drive shaft — 24 pulses/revolution, providing precise row-advance detection'),
          twoColRow('Row detection', 'One full encoder tick = 5 mm card advance = one row read. The Arduino fires sensor reading on each encoder tick.', true),
          twoColRow('Card loop', 'Optional: small take-up spool or tray to receive the card as it exits. For loop playback, tape the card ends together.'),
        ]
      }),
      body(''),

      h3('Electronics'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Microcontroller', 'Arduino Uno or Nano — reads encoder, polls 24 sensors, sends data over USB serial or MIDI'),
          twoColRow('Multiplexer', '3 × CD74HC4051 8-channel analog multiplexers — allows 24 sensors to be read on 3 Arduino analog pins', true),
          twoColRow('Pull-up resistors', '24 × 10 kΩ resistors for phototransistor voltage dividers'),
          twoColRow('MIDI output', 'Arduino with MIDI-over-USB library (or hardware DIN-5 MIDI out for older setups)', true),
          twoColRow('Power', '5V USB from the laptop — no separate power supply needed'),
        ]
      }),
      body(''),

      h2('Software Architecture'),

      h3('Arduino firmware'),
      body('The Arduino sketch does three things: detects encoder ticks, reads all 24 sensors via the multiplexers on each tick, and transmits the 24-bit reading to the laptop. The transmission format is simple: a MIDI SysEx message containing the 24 sensor values (0 or 1), or a simple serial byte stream that SuperCollider reads directly.'),
      body('Debouncing the encoder is important — use the standard Encoder library, which handles this robustly. Sensor thresholding: read each sensor as a raw analog value (0–1023) and apply a threshold (typically around 500) to classify as open or closed. Calibration: on startup, read all sensors with no card inserted to establish the baseline "open" reading.', { spAfter: 240 }),

      h3('SuperCollider patch'),
      body('SuperCollider receives the 24-bit row reading and maps it to a three-engine synthesis architecture. The card is divided into three horizontal regions, each controlling a different synthesis engine:'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 3280, 3280],
        rows: [
          tableHeader3('Region', 'Synthesis engine', 'Sonic character'),
          threeColRow('Top 8 rows', 'Additive / spectral', 'Bright shimmer, harmonic overtones', false),
          threeColRow('Middle 8 rows', 'Formant / vocal', 'Vowel resonances, body', true),
          threeColRow('Bottom 8 rows', 'Karplus-Strong strings', 'Plucked/bowed resonance', false),
        ]
      }),
      body(''),
      body('Within each region, the column position of each hole controls the spatial and timbral character of the voice it activates. Left columns produce brighter, more forward sounds; right columns produce darker, more recessive ones. Stereo pan follows column position. The density of holes in any region controls the overall loudness and harmonic complexity of that engine.'),
      body('The crank speed is directly perceptible in the music. Cranking fast produces rapid parameter changes — a feeling of the sound rushing forward. Cranking slowly produces long, sustained readings of each row — meditative, static, evolving only as the pattern changes. Stopping altogether freezes the sound at whatever row the card rests on.', { spAfter: 240 }),

      h2('Punch Card → Sound Mappings'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          tableHeader2('Card feature', 'Sonic result'),
          mappingRow('Single isolated hole', 'A clear bell tone or harmonic partial that decays slowly', false),
          mappingRow('Dense horizontal row', 'Thick harmonic cluster, all voices active', true),
          mappingRow('Empty row', 'Silence — or a very quiet residual drone', false),
          mappingRow('Holes in top region only', 'Bright, crystalline, metallic — high harmonics', true),
          mappingRow('Holes in bottom region only', 'Deep, physical, warm — bass resonance and string', false),
          mappingRow('Holes scattered mid-region', 'Vowel formants — the sound acquires voice-like quality', true),
          mappingRow('Symmetric pattern (mirror-XY)', 'Stereo balance, consonant intervals, stable feeling', false),
          mappingRow('Asymmetric scatter', 'Incoherent partials, rough, tense', true),
          mappingRow('Diagonal pattern (top-left to bottom-right)', 'Descending sweep — bright to dark as card advances', false),
          mappingRow('Dense cluster at edges', 'High FM index — metallic, industrial, distorted', true),
          mappingRow('Vertical repeat (same row several times)', 'Loop — the pattern sustains and pulses', false),
          mappingRow('Gradual density increase', 'Crescendo — sound builds as more voices activate', true),
        ]
      }),
      body(''),

      h2('Bill of Materials — Project 1'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 2200, 2480, 2480],
        rows: [
          tableHeader4([['Component', 2200], ['Quantity', 2200], ['Source', 2480], ['Approx. cost', 2480]]),
          fourColRow('Arduino Nano', '1', 'Arduino.cc / AliExpress', '€5–25', false),
          fourColRow('LTR-301 phototransistors', '24 + spares', 'Mouser / LCSC', '€0.15 each', true),
          fourColRow('3mm warm-white LEDs', '24 + spares', 'Any electronics supplier', '€0.05 each', false),
          fourColRow('CD74HC4051 multiplexer', '3', 'Mouser / LCSC', '€0.50 each', true),
          fourColRow('10 kΩ resistors', '30 (pack)', 'Any electronics supplier', '€1', false),
          fourColRow('Bourns PEC11R encoder', '1', 'Mouser', '€3', true),
          fourColRow('Rubber O-rings (rollers)', '4 × M8', 'Hardware store', '€2', false),
          fourColRow('Aluminium rod 8mm', '200 mm', 'Hardware store / online', '€3', true),
          fourColRow('Laser-cut acrylic/plywood', '1 set', 'Local maker space', '€10–20', false),
          fourColRow('Misc: wire, connectors, screws', '—', 'Any supplier', '€10', true),
          fourColRow('Laptop running SuperCollider', '1', 'Existing', '—', false),
          fourColRow('Small powered speakers', '1 pair', 'Existing / secondhand', '€20–60', true),
          fourColRow('', '', 'Total estimate', '€60–120', false),
        ]
      }),
      body(''),

      h2('Build Timeline — Project 1'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Week 1', 'Order components. Design enclosure in CAD (Fusion 360 or Inkscape for laser cutting). Build breadboard prototype with 3 sensors and the encoder to validate the circuit.'),
          twoColRow('Week 2', 'Assemble full 24-sensor array. Write Arduino firmware. Validate serial output with a simple Python script.', true),
          twoColRow('Week 3', 'Write SuperCollider patch. Map sensor readings to synthesis parameters. Test with a real punch card.'),
          twoColRow('Week 4', 'Build final enclosure. Calibrate sensor thresholds. Run first workshop test session. Refine sound design based on feedback.', true),
          twoColRow('Ongoing', 'Iterate on SuperCollider patch. Develop a library of punch card patterns with known sonic characters. Document for participants.'),
        ]
      }),
      body(''),
      new Paragraph({ children: [new PageBreak()] }),

      // ══════════════════════════════════════════════════════════
      // PROJECT 2
      // ══════════════════════════════════════════════════════════
      h1('Project 2: The Singing Knitting Machine'),
      body('A vintage knitting machine — Brother or Silver Reed — modified to emit ambient sound as it knits. The punch card feeds normally. The carriage moves. A scarf slowly emerges. Simultaneously, speakers emit drones and harmonics derived from the pattern passing through the machine. The machine knits music. Or perhaps: it remembers through fabric and sound simultaneously.'),
      body('This is the more ambitious and more conceptually powerful object. It requires more patience to build but no more technical knowledge. And the spectacle — a machine knitting while sound emerges from it — is the kind of thing people gather around.', { spAfter: 240 }),

      h2('Two Approaches'),

      h3('Approach A: Pre-read the card optically (recommended starting point)'),
      body('Install a sensor strip — identical to Project 1 — just before the punch card enters the knitting mechanism. The card is read optically before it reaches the selector pins. This approach is completely non-destructive: the machine continues to function exactly as intended, and the reading happens to the card before it enters the mechanical system.'),
      body('The sensor fires on each carriage pass rather than on an encoder tick. A reed switch or optical interrupter on the carriage detects direction and position. Each time the carriage completes a pass, the Arduino reads the current card row and sends it to SuperCollider. Since the card advances one row per pass, the music and the fabric are always synchronised — they are reading the same row at the same moment.'),
      body('Advantages: non-destructive, easy to install, fully reversible. The machine can be returned to normal operation in minutes. This is the right starting point for a first build.', { spAfter: 240 }),

      h3('Approach B: Read the selector mechanism itself (advanced)'),
      body('The knitting machine\'s selector pins respond to the card\'s holes by physically moving to engage or disengage needles. This mechanical response — the actual decisions the machine makes — can be read with Hall effect sensors, optical interrupters, or microswitches placed near the selector bed.'),
      body('This approach produces a fundamentally different artistic result. The music is not what the card says. It is what the machine does. If the machine misreads a hole — a common occurrence with worn cards or dirty pins — that misread becomes part of the composition. A yarn jam creates silence. The tension in the carriage becomes tension in the sound. The machine\'s imperfections and character are audible.'),
      body('This has a strong experimental music precedent: the idea that the instrument\'s own behavior, including its failures, is part of the composition. It is closer to David Tudor\'s work with unpredictable circuits, or the aesthetic of early industrial and noise music, than to conventional instrument design. For the right context and audience it is the more interesting build. For a first workshop, start with Approach A.', { spAfter: 240 }),

      h2('Hardware — Approach A'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Sensor strip', 'Identical to Project 1: 24 × phototransistors + 24 × LEDs, mounted in a printed or machined bracket that attaches to the card guide just before the selector. The bracket clips on without modification to the machine.'),
          twoColRow('Carriage detection', 'A small neodymium magnet glued to the carriage + a reed switch fixed to the needle bed frame. Each carriage pass (left and right) triggers a read.', true),
          twoColRow('Microcontroller', 'Same Arduino Nano as Project 1 — reads the reed switch interrupt, polls sensors, transmits over USB serial.'),
          twoColRow('Speakers', 'Two small full-range drivers (50–80 mm) built into the machine\'s base or a separate enclosure positioned beside it. Powered by a small class-D amplifier board (PAM8403 or similar, €3–5).', true),
          twoColRow('Optional: projector', 'A laptop running Processing or TouchDesigner projects the current card row enlarged on a wall or screen above the machine — participants can see exactly which row is being read.'),
        ]
      }),
      body(''),

      h2('Synchronisation'),
      body('The fundamental challenge in Project 2 is keeping the music and the fabric synchronised. In Project 1, the crank controls both: one encoder tick = one row advance = one new reading. In Project 2, the carriage controls both: one pass = one row advance = one new reading.'),
      body('The carriage on a standard Brother/Silver Reed machine makes one pass per row of knitting. The speed of the pass is controlled by the knitter — a slow careful pass, or a quick confident one. This variation is audible in the music: a slow pass produces a sustained reading of that row; a quick pass produces a brief flash of its sonic character before the next row takes over.'),
      body('This is an important artistic decision: the knitter is the performer. Their physical engagement with the machine — the speed, rhythm, and hesitation of each pass — shapes the music as directly as the pattern does.', { spAfter: 240 }),

      h2('Sound Design for Project 2'),
      body('The SuperCollider patch for Project 2 can be identical to Project 1 at the synthesis level. The difference is in the musical structure that emerges from the knitting process itself.'),
      body('A typical garment involves hundreds of rows. At one pass every two to four seconds, that is five to ten minutes of music for a short section of a scarf. The music evolves slowly — more slowly than in Project 1, where the crank speed is unconstrained. This slower evolution is appropriate for the context: knitting is a slow practice, and music that unfolds at the same pace as the fabric feels right.'),
      body('Repeating patterns — the whole point of a punch card — produce musical loops. If the card is 24 rows long and loops continuously, the music repeats on the same cycle as the pattern in the fabric. The texture builds as row after row of the same pattern accumulates in the growing fabric. This is genuinely beautiful: the physical accumulation of the garment is audible as a deepening and thickening of the sound.', { spAfter: 240 }),

      h2('Build Timeline — Project 2'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Week 1–2', 'Acquire a working Brother or Silver Reed machine (secondhand, €30–150). Study its card feed mechanism. Design and 3D print or laser-cut the sensor bracket.'),
          twoColRow('Week 3', 'Install sensor strip and carriage reed switch. Wire to Arduino. Validate readings with a test card and a serial monitor.', true),
          twoColRow('Week 4', 'Adapt the SuperCollider patch from Project 1. Test a full knitting session — at least 60 rows — and record the audio. Adjust sound design.'),
          twoColRow('Week 5–6', 'Install speakers and amplifier. Refine cabling and mounting. Run a test workshop session with 2–3 knitters.', true),
          twoColRow('Ongoing', 'Develop a library of cards with known sonic characters. Explore Approach B as an optional parallel build for more experimental contexts.'),
        ]
      }),
      body(''),
      new Paragraph({ children: [new PageBreak()] }),

      // ══════════════════════════════════════════════════════════
      // WORKSHOP STRUCTURE
      // ══════════════════════════════════════════════════════════
      h1('Workshop Structure'),
      body('The two builds support different workshop formats. Both can be run independently or together. The following describes the full format combining both instruments.', { spAfter: 240 }),

      h2('The Room'),
      body('The knitting machine sits at the centre of the space. The punch card synth (Project 1) sits on a separate table where participants can interact with it directly during the design phase. A projector displays the current card row enlarged on the wall. Speakers are positioned so the sound fills the room rather than localising to the machines.'),
      body('The room should feel like a studio, not a classroom. Materials available throughout: yarn in several colours, blank punch cards, hole punches, reference patterns, printed examples of patterns and their sonic characters.', { spAfter: 240 }),

      h2('Session Arc — One Evening (3–4 hours)'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Opening (30 min)', 'Introduction to the project. Brief listening exercise: play several cards through the punch card synth without showing participants the patterns first. Ask them to describe what they hear. Then reveal the cards. Discuss the correspondence between visual pattern and sonic character.'),
          twoColRow('Design phase (60 min)', 'Participants design their own punch cards using the online tool (jordanelias.de/blog/punch-card-music/). They hear each design through the browser synth as they draw, then print their card. Facilitated discussion: what sonic character are you aiming for? What does your pattern look like? Does it sound like you expected?', true),
          twoColRow('Machine session (90 min)', 'Cards are fed through the knitting machine one at a time. Each participant\'s card is inserted, the machine begins knitting, and the room listens. The projector shows the current row. The growing fabric is visible. This continues until each participant\'s card has been run at least once — or the group chooses to continue with a favourite card.'),
          twoColRow('Listening and reflection (30 min)', 'The accumulated fabric hangs or lays across the space. Recordings from the session play back quietly. Cards are displayed alongside their sections of fabric. Participants share observations. What surprised them? What did their pattern sound like? What would they change?', true),
          twoColRow('Closing (30 min)', 'Each participant leaves with their section of fabric and a printed copy of their punch card. A recording of the full session is available to share.'),
        ]
      }),
      body(''),

      h2('Facilitation Notes'),
      bullet('The machine session works best with one person designated as the knitter — someone comfortable with the machine who can keep a steady pace. The role can rotate.'),
      bullet('Encourage participants to knit slowly during the first pass of their card, then experiment with different speeds. The change in musical character is striking and immediate.'),
      bullet('If something goes wrong — a yarn jam, a dropped stitch, a misread card — treat it as part of the composition. Narrate it as such. This reframe is important and usually lands well.'),
      bullet('The projector display is more important than it might seem: seeing the row enlarged makes the connection between pattern and sound explicit in a way that looking at the small card does not.'),
      bullet('For groups new to experimental music: a brief listening exercise at the start is essential. Play two minutes of Éliane Radigue or Pauline Oliveros. Not to explain it, but to calibrate the room to the idea that sound does not need to be melodic to be musical.'),
      body(''),

      h2('What the Evening Produces'),
      body('By the end of the session the group will have made:'),
      bullet('A collective scarf or fabric panel — each participant\'s card producing a distinct section, all knitted from the same yarn'),
      bullet('A recording of the full musical performance — one continuous piece composed collectively by the group\'s cards'),
      bullet('A set of printed punch cards, displayable alongside the fabric they produced'),
      bullet('A shared experience of the connection between visual pattern, physical making, and sound'),
      body('These four outputs can be exhibited together. The fabric, the cards, a recording playing quietly, and perhaps a screen showing the online tool. The workshop becomes an installation.', { spAfter: 240 }),

      new Paragraph({ children: [new PageBreak()] }),

      // ══════════════════════════════════════════════════════════
      // TECHNICAL APPENDIX
      // ══════════════════════════════════════════════════════════
      h1('Technical Appendix'),

      h2('Arduino Firmware Outline'),
      body('The following pseudocode describes the core logic of the Arduino sketch for both projects. The actual implementation uses the Encoder library for the rotary encoder and the MIDIUSB library for MIDI over USB.', { spAfter: 120 }),
      new Paragraph({
        ...sp(80, 80),
        indent: { left: 560 },
        children: [new TextRun({
          text: [
            'on encoder tick (Project 1) or carriage pass (Project 2):',
            '    for each sensor group 0..2:',
            '        select channel on multiplexer',
            '        read 8 analog values',
            '        threshold each to 0 or 1',
            '    assemble 24-bit reading',
            '    send as MIDI SysEx or serial bytes',
            '    advance row counter',
            '',
            'calibration on startup:',
            '    read all sensors with no card',
            '    store as "open" baseline',
            '    threshold = baseline × 0.6',
          ].join('\n'),
          font: 'Courier New',
          size: 18,
          color: C.darkGrey,
        })]
      }),
      body(''),

      h2('SuperCollider Patch Architecture'),
      body('The SuperCollider patch receives serial or MIDI data and maintains a 24-element array representing the current card row. This array is divided into three regions and used to control three synthesis engines running in parallel.'),
      body('Engine A (additive): a single PeriodicWave oscillator rebuilt each time the top-region array changes. Each hole contributes a sine partial at a specific harmonic (1×, 2×, 3×... up to 32×). Phase of each partial is set by column position. The wave is rebuilt atomically using Buffer.alloc and the waveform is swapped without clicking.'),
      body('Engine B (formant): three bandpass filters driven by a sawtooth oscillator at the root frequency. Middle-region holes shift formant frequencies (F1: 200–900 Hz, F2: 700–2500 Hz, F3: 1800–3500 Hz) and Q values. Specific hole configurations produce recognisable vowel sounds.'),
      body('Engine C (strings): Karplus-Strong physical string synthesis. Each bottom-region hole activates a string voice. Row position sets pitch via delay line length. Column position sets sustain via the loop filter frequency. Voices are stereo-panned by column. All three engines share a root frequency and feed into a master bus with reverb.', { spAfter: 240 }),

      h2('Punch Card Specification Reference'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 6560],
        rows: [
          twoColRow('Column count', '24 (standard Brother/Silver Reed/Toyota domestic machines)'),
          twoColRow('Hole spacing', '5 mm centre-to-centre, both horizontally and vertically', true),
          twoColRow('Hole diameter', '4–4.5 mm (standard punch)'),
          twoColRow('Card material', 'Stiff paper or mylar — mylar is more durable for repeated use', true),
          twoColRow('Card length', 'Variable — typically 24, 40, or 60 rows for standard patterns'),
          twoColRow('Compatible machines', 'Brother KH-860, KH-940, KH-970; Silver Reed SK-280, LK-140; Toyota K747 and most domestic Japanese-made machines from 1965–1995', true),
          twoColRow('Printing/punching', 'Export from the online tool at jordanelias.de/blog/punch-card-music/ as a print-ready PDF at 100% scale, then punch with a standard 4.5mm paper punch or use a pre-punched card as a template'),
        ]
      }),
      body(''),

      h2('Sourcing a Knitting Machine'),
      body('Vintage domestic knitting machines in working condition are readily available secondhand throughout Europe. Brother and Silver Reed machines are the most common and best-documented. Look for:'),
      bullet('Brother KH-860 or KH-940 — widely available, well-documented, spare parts available online'),
      bullet('Silver Reed SK-280 — slightly less common but excellent build quality'),
      bullet('Price range: €30–150 on eBay, Kleinanzeigen, Vinted, or local charity shops'),
      bullet('Essential accessories: the ribber is not needed; the main bed and card reader are sufficient'),
      bullet('Check that the card reader mechanism is present and undamaged — this is the part most often missing on incomplete machines'),
      body('Test the machine before modifying it. Knit a 20-row swatch by hand (without a card) to confirm the needle bed, yarn feeder, and carriage are functioning. A working machine is far easier to modify than a broken one.', { spAfter: 240 }),

      h2('Further Resources'),
      bullet('SuperCollider documentation: supercollider.github.io'),
      bullet('Arduino Encoder library: github.com/PaulStoffregen/Encoder'),
      bullet('Punch card pattern tool: jordanelias.de/blog/punch-card-music/'),
      bullet('Knitting machine repair guides: machine-knitting.uk and YouTube channels by Diana Sullivan and Alessandrina Lazzaro'),
      bullet('Karplus-Strong synthesis in SuperCollider: "The SuperCollider Book" (MIT Press, 2011), Chapter 11'),
      bullet('Pauline Oliveros, Deep Listening — essential reading for the acoustic philosophy underpinning this work'),
      bullet('Éliane Radigue — her ADNOS trilogy and Trilogie de la Mort as reference points for drone and textural composition'),
      body(''),
      rule(),
      body('Document prepared by Jordan Elias, MT-BC · Berlin · 2026 · jordanelias.de', {
        align: AlignmentType.CENTER, color: C.midGrey, size: 18, spAfter: 80
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/textile-sound-workshop-plan.docx', buffer);
  console.log('Done');
});
