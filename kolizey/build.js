const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const gi = require("react-icons/gi");
const fa = require("react-icons/fa6");

const OUT_DIR = "/mnt/project-files/kolizey";
const PHOTO_DIR = path.join(OUT_DIR, "photos");
const OUT = path.join(OUT_DIR, "Презентация_Колизей.pptx");

const DARK = "3B2418", TERRA = "B5532F", GOLD = "E0AE4F", STONE = "EFE4D3", INK = "2B1D14", MUTED = "6B5646", WHITE = "FFFFFF";
const HEAD = "Cambria", BODY = "Calibri";

async function icon(Comp, color, px = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: px }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function findPhoto(key) {
  for (const ext of [".jpg", ".jpeg", ".png", ".webp"]) {
    const p = path.join(PHOTO_DIR, key + ext);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

let CAMERA;
// Photo slot: real photo cropped to fit if photos/<key>.* exists, otherwise a marked placeholder.
async function photo(slide, key, label, x, y, w, h) {
  const p = findPhoto(key);
  if (p) {
    const pxW = 1600, pxH = Math.round((1600 * h) / w);
    const buf = await sharp(p).rotate().resize(pxW, pxH, { fit: "cover", position: "attention" }).jpeg({ quality: 85 }).toBuffer();
    slide.addImage({ data: "image/jpeg;base64," + buf.toString("base64"), x, y, w, h, rounding: false, altText: label, objectName: "Фото: " + label,
      shadow: { type: "outer", color: "000000", opacity: 0.25, blur: 6, offset: 2, angle: 90 } });
    return;
  }
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.12, fill: { color: STONE }, line: { color: TERRA, width: 1.5, dashType: "dash" }, objectName: "Место для фото: " + label });
  const s = Math.min(0.6, h * 0.3);
  slide.addImage({ data: CAMERA, x: x + w / 2 - s / 2, y: y + h / 2 - s / 2 - 0.2, w: s, h: s });
  slide.addText("Фото: " + label, { x: x + 0.1, y: y + h / 2 + s / 2 - 0.1, w: w - 0.2, h: 0.45, fontFace: BODY, fontSize: 12, color: MUTED, align: "center", italic: true, isTextBox: true });
}

(async () => {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  CAMERA = await icon(fa.FaCamera, TERRA);
  const COL_GOLD = await icon(gi.GiColiseum, GOLD);
  const COL_TERRA = await icon(gi.GiColiseum, TERRA);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = "Максимов Илья";
  pres.title = "Колизей — объект Всемирного наследия";

  const title = (slide, text, color = INK) =>
    slide.addText(text, { x: 0.5, y: 0.3, w: 9, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true, color, margin: 0, isTextBox: true });

  // 1. Титульный
  {
    const s = pres.addSlide();
    s.background = { color: DARK };
    s.addImage({ data: COL_GOLD, x: 0.5, y: 0.55, w: 0.8, h: 0.8 });
    s.addText("Окружающий мир", { x: 0.5, y: 1.45, w: 4.5, h: 0.4, fontFace: BODY, fontSize: 16, color: GOLD, margin: 0, isTextBox: true });
    s.addText("Колизей", { x: 0.5, y: 1.85, w: 4.5, h: 1.0, fontFace: HEAD, fontSize: 54, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText("объект Всемирного наследия", { x: 0.5, y: 2.85, w: 4.5, h: 0.5, fontFace: HEAD, fontSize: 24, italic: true, color: STONE, margin: 0, isTextBox: true });
    s.addText([
      { text: "Выполнил: ", options: { color: GOLD } },
      { text: "Максимов Илья", options: { bold: true, breakLine: true } },
      { text: "ученик 4 «Г» класса" },
    ], { x: 0.5, y: 3.95, w: 4.5, h: 0.8, fontFace: BODY, fontSize: 16, color: WHITE, margin: 0, isTextBox: true });
    await photo(s, "01_snaruzhi", "Колизей снаружи", 5.3, 0.5, 4.2, 4.6);
  }

  // 2. Что такое Колизей
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    title(s, "Что такое Колизей?");
    await photo(s, "02_tselikom", "Колизей целиком", 0.5, 1.3, 4.3, 3.8);
    s.addText([
      { text: "Колизей — огромный древний цирк-амфитеатр в городе Риме (Италия). Его построили почти 2000 лет назад, в 80 году.", options: { breakLine: true } },
      { text: "Основал его император Веспасиан, а достроил его сын Тит.", options: { bold: true, color: TERRA } },
    ],
      { x: 5.2, y: 1.2, w: 4.3, h: 1.45, fontFace: BODY, fontSize: 14, paraSpaceAfter: 4, color: INK, margin: 0, valign: "top", isTextBox: true });
    const stats = [["50 000", "зрителей помещалось"], ["48 м", "высота, как дом в 16 этажей"], ["80", "входов и выходов"], ["8 лет", "шла стройка"]];
    stats.forEach(([n, l], i) => {
      const cx = 5.2 + (i % 2) * 2.2, cy = 2.75 + Math.floor(i / 2) * 1.3;
      s.addShape("roundRect", { x: cx, y: cy, w: 2.05, h: 1.2, rectRadius: 0.1, fill: { color: STONE }, line: { type: "none" } });
      s.addText(n, { x: cx, y: cy + 0.1, w: 2.05, h: 0.55, fontFace: HEAD, fontSize: 26, bold: true, color: TERRA, align: "center", margin: 0, isTextBox: true });
      s.addText(l, { x: cx + 0.1, y: cy + 0.65, w: 1.85, h: 0.45, fontFace: BODY, fontSize: 12, color: INK, align: "center", margin: 0, isTextBox: true });
    });
  }

  // 3. Что здесь происходило
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    title(s, "Что здесь происходило?");
    const cols = [
      ["03_gladiatory", "гладиаторы", "Бои гладиаторов", "Сильные воины сражались на арене, а зрители болели за них."],
      ["04_arena", "арена внутри", "Дикие звери", "Из далёких стран привозили львов, медведей и слонов."],
      ["05_podzemelya", "подземелья под ареной", "Подземелья", "Под ареной были клетки и подъёмники для зверей."],
    ];
    for (let i = 0; i < 3; i++) {
      const [key, lbl, h, t] = cols[i];
      const x = 0.5 + i * 3.1;
      await photo(s, key, lbl, x, 1.25, 2.8, 2.35);
      s.addText(h, { x, y: 3.75, w: 2.8, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: TERRA, margin: 0, isTextBox: true });
      s.addText(t, { x, y: 4.15, w: 2.8, h: 0.9, fontFace: BODY, fontSize: 14, color: INK, margin: 0, valign: "top", isTextBox: true });
    }
  }

  // 4. Почему это всемирное наследие
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    title(s, "Почему Колизей под охраной?");
    const rows = [
      [fa.FaGlobe, "ЮНЕСКО", "Эта организация берёт под охрану самые ценные места Земли."],
      [fa.FaLandmark, "1980 год", "Центр Рима вместе с Колизеем внесли в список Всемирного наследия."],
      [gi.GiColiseum, "2007 год", "Колизей выбрали одним из Новых семи чудес света."],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [Ic, h, t] = rows[i];
      const y = 1.3 + i * 1.25;
      s.addShape("ellipse", { x: 0.5, y, w: 0.8, h: 0.8, fill: { color: TERRA }, line: { type: "none" } });
      s.addImage({ data: await icon(Ic, WHITE), x: 0.68, y: y + 0.18, w: 0.44, h: 0.44 });
      s.addText(h, { x: 1.5, y: y - 0.05, w: 3.4, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: TERRA, margin: 0, isTextBox: true });
      s.addText(t, { x: 1.5, y: y + 0.35, w: 3.4, h: 0.75, fontFace: BODY, fontSize: 14, color: INK, margin: 0, valign: "top", isTextBox: true });
    }
    await photo(s, "06_rim", "Рим вокруг Колизея", 5.3, 1.25, 4.2, 2.2);
    await photo(s, "07_unesco", "знак ЮНЕСКО", 5.3, 3.6, 2.0, 1.5);
    await photo(s, "08_vnutri", "Колизей изнутри", 7.5, 3.6, 2.0, 1.5);
  }

  // 5. Колизей сегодня
  {
    const s = pres.addSlide();
    s.background = { color: DARK };
    title(s, "Колизей сегодня", WHITE);
    await photo(s, "09_noch", "Колизей ночью", 0.5, 1.25, 4.4, 2.5);
    await photo(s, "10_turisty", "туристы у Колизея", 5.1, 1.25, 4.4, 2.5);
    s.addText([
      { text: "Каждый год сюда приезжают миллионы туристов.", options: { bullet: true, breakLine: true } },
      { text: "Учёные укрепляют старые стены и чинят арену.", options: { bullet: true, breakLine: true } },
      { text: "Царапать стены и уносить камни строго запрещено.", options: { bullet: true } },
    ], { x: 0.5, y: 3.95, w: 5.6, h: 1.2, fontFace: BODY, fontSize: 14, color: WHITE, paraSpaceAfter: 4, margin: 0, valign: "top", isTextBox: true });
    s.addImage({ data: COL_GOLD, x: 6.35, y: 4.05, w: 0.55, h: 0.55 });
    s.addText("Берегите чудеса света, чтобы их увидели и наши внуки!", { x: 7.0, y: 3.95, w: 2.5, h: 1.0, fontFace: HEAD, fontSize: 15, italic: true, bold: true, color: GOLD, margin: 0, valign: "top", isTextBox: true });
  }

  pres._slides[pres._slides.length - 1].addNotes("Фото: Wikimedia Commons\n01: Colosseum-Rome-May-2009.jpg (Acediscovery, CC BY 4.0)\n02: Colosseum of Rome and Roman forum.jpg (Wilfredor, CC0)\n03: Gladiators from the Zliten mosaic 3 cropped.JPG (мозаика из Злитена, общественное достояние)\n04: Gladiators face animals Zliten mosaic.JPG (мозаика из Злитена, общественное достояние)\n05: Hypogeum 1 (15005526662).jpg (daryl_mitchell, CC BY-SA 2.0)\n06: Roman Forum and the Colosseum.jpg (Michael Wilson, CC BY 2.0)\n07: World Heritage Logo global.svg (эмблема ЮНЕСКО, общественное достояние)\n08: Colosseum Arena (5986632567).jpg (Erik Drost, CC BY 2.0)\n09: Colosseum at night - wide angle.JPG (Martin Falbisoner, CC BY-SA 3.0)\n10: Colosseum Interior 1 (15005911295).jpg (daryl_mitchell, CC BY-SA 2.0)");
  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
