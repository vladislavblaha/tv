// xml_do_xlsx.js - prevod JMHZ XML (CSSZ) rovnou do Excelu (.xlsx)
// Vyzaduje nainstalovany Microsoft Excel. Windows Script Host (JScript).
// Pouziti: pretahnete XML soubor mysi na tento skript, nebo
//          dejte skript a XML do jedne slozky a dvakrat kliknete na skript.

function trim(s) { return String(s).replace(/^\s+|\s+$/g, ""); }
function parseNum(s) {
  var t = trim(s);
  if (t === "") return 0;
  var n = parseFloat(t);
  return isNaN(n) ? 0 : n;
}
function colLetter(n) { return String.fromCharCode(64 + n); }
function rgb(r, g, b) { return r + g * 256 + b * 65536; }

var fso = new ActiveXObject("Scripting.FileSystemObject");
var xmlPath = "";

if (WScript.Arguments.length > 0) {
  xmlPath = WScript.Arguments(0);
} else {
  var folder = fso.GetParentFolderName(WScript.ScriptFullName);
  var files = new Enumerator(fso.GetFolder(folder).Files);
  for (; !files.atEnd(); files.moveNext()) {
    var fl = files.item();
    var nm = fl.Name.toLowerCase();
    if (nm.substring(0, 5) === "jmhz_" && nm.substring(nm.length - 4) === ".xml") {
      xmlPath = fl.Path;
      break;
    }
  }
  if (xmlPath === "") {
    WScript.Echo("Nena\u0161el jsem soubor jmhz_*.xml ve slo\u017Ece skriptu.\nP\u0159et\u00E1hn\u011Bte XML soubor my\u0161\u00ED na tento skript.");
    WScript.Quit(1);
  }
}
xmlPath = fso.GetAbsolutePathName(xmlPath);

var doc = new ActiveXObject("MSXML2.DOMDocument.6.0");
doc.async = false;
doc.validateOnParse = false;
doc.resolveExternals = false;
doc.setProperty("SelectionLanguage", "XPath");
doc.setProperty("SelectionNamespaces",
  "xmlns:n1='http://schemas.cssz.cz/JMHZ/podani/1.0' " +
  "xmlns:form='http://schemas.cssz.cz/JMHZ/form/1.0' " +
  "xmlns:so='http://schemas.cssz.cz/JMHZ/souhrn/1.0' " +
  "xmlns:pvpoj='http://schemas.cssz.cz/JMHZ/PVPOJ/1.0'");
if (!doc.load(xmlPath)) {
  WScript.Echo("XML se nepoda\u0159ilo na\u010D\u00EDst:\n" + doc.parseError.reason);
  WScript.Quit(1);
}

var paths = [
  "form:souhrnDataZec/form:prijmy/form:zuctovanoCelkem",
  "form:souhrnDataZec/form:prijmy/form:osvobozenoCelkem",
  "form:souhrnDataZec/form:zalohaNaDan/form:zakladDane",
  "form:souhrnDataZec/form:zalohaNaDan/form:vypoctenaZaloha",
  "form:souhrnDataZec/form:zalohaNaDan/form:danZalohaPoSleve",
  "form:souhrnDataZec/form:zvlastniSazbaDane/form:srazenaDan",
  "form:souhrnDataZec/form:mzdaCista/form:mzdaCista",
  "form:souhrnDataZec/form:zdravPojZamestnanec/form:zdravotniPojisteni",
  "form:souhrnDataZec/form:zdravPojZamestnavatel/form:zdravotniPojisteni",
  "form:pojisteni/form:pojisteniZamestnanec/form:socialniPojisteni",
  "form:pojisteni/form:pojisteniZamestnavatel/form:socialniPojisteni",
  "form:pojisteni/form:vymerovaciZakladParagraf5/form:pismenoA",
  "form:prubehZamestnani/form:odpracovaneHodiny/form:pocet",
  "form:prubehZamestnani/form:odpracovaneHodiny/form:rozpad/form:prescas"
];

var hdr = ["Os. \u010D\u00EDslo", "Jm\u00E9no", "Pom\u011Br", "Typ", "IK MPSV", "Z\u00FA\u010Dtovan\u00FD p\u0159\u00EDjem",
  "Osvobozen\u00E9 p\u0159\u00EDjmy", "Z\u00E1klad z\u00E1lohy na da\u0148", "Vypo\u010Dten\u00E1 z\u00E1loha", "Z\u00E1loha po slev\u00E1ch",
  "Da\u0148 zvl. sazba", "\u010Cist\u00FD p\u0159\u00EDjem", "ZP zam\u011Bstnanec", "ZP zam\u011Bstnavatel",
  "SP zam\u011Bstnanec", "SP zam\u011Bstnavatel", "VZ pro zam\u011Bstnavatele (\u00A75a p\u00EDsm. a)",
  "Odpracovan\u00E9 hodiny", "Z toho p\u0159es\u010Das"];

var xl = null;
try {
  xl = new ActiveXObject("Excel.Application");
} catch (e) {
  WScript.Echo("Nepoda\u0159ilo se spustit Excel. Je nainstalovan\u00FD?\n" + e.message);
  WScript.Quit(1);
}

try {
  xl.Visible = true;
  xl.DisplayAlerts = false;
  var wb = xl.Workbooks.Add();
  var ws = wb.Worksheets(1);
  ws.Name = "Zam\u011Bstnanci";

  var c, k;
  for (c = 0; c < hdr.length; c++) ws.Cells(1, c + 1).Value = hdr[c];

  var forms = doc.selectNodes("//n1:formularOsoby");
  var cnt = 0, idx, r = 1;
  for (idx = 0; idx < forms.length; idx++) {
    var fo = forms.item(idx);
    r++;

    var osc = "", jmeno = "", pomer = "";
    var p = fo.previousSibling;
    while (p !== null && p.nodeType !== 8 && p.nodeType !== 1) p = p.previousSibling;
    if (p !== null && p.nodeType === 8) {
      var comm = trim(p.nodeValue).replace(/\s+/g, " ");
      var i1 = comm.indexOf(" ");
      var i2 = comm.lastIndexOf(" ");
      if (i1 > 0 && i2 > i1) {
        osc = comm.substring(0, i1);
        jmeno = comm.substring(i1 + 1, i2);
        pomer = comm.substring(i2 + 1);
      } else {
        jmeno = comm;
      }
    }

    var typ = "Hlavn\u00ED";
    var b = fo.selectSingleNode("form:bezPriznaku");
    if (b === null) {
      b = fo.selectSingleNode("form:cinnostKS");
      typ = "\u010Cinnost KS";
    }
    var ik = "", v;
    if (b !== null) {
      v = b.selectSingleNode("form:identifikace/form:ikMpsv");
      if (v !== null) ik = v.text;
    }

    ws.Cells(r, 1).NumberFormat = "@";
    ws.Cells(r, 1).Value = osc;
    ws.Cells(r, 2).Value = jmeno;
    ws.Cells(r, 3).NumberFormat = "@";
    ws.Cells(r, 3).Value = pomer;
    ws.Cells(r, 4).Value = typ;
    ws.Cells(r, 5).NumberFormat = "@";
    ws.Cells(r, 5).Value = ik;

    for (k = 0; k < paths.length; k++) {
      var n = 0;
      if (b !== null) {
        v = b.selectSingleNode(paths[k]);
        if (v !== null) n = parseNum(v.text);
        if (k === 0 && typ === "\u010Cinnost KS") {
          v = b.selectSingleNode("form:prijem/form:dan/form:zakladDane");
          if (v !== null) n = parseNum(v.text);
        }
      }
      ws.Cells(r, 6 + k).Value = n;
    }
    if (typ === "\u010Cinnost KS") {
      ws.Range("A" + r + ":S" + r).Interior.Color = rgb(255, 242, 204);
    }
    cnt++;
  }

  var last = r;
  var tot = last + 1;
  ws.Cells(tot, 2).Value = "CELKEM";
  for (k = 0; k < paths.length; k++) {
    var L = colLetter(6 + k);
    ws.Cells(tot, 6 + k).Formula = "=SUM(" + L + "2:" + L + last + ")";
  }
  ws.Range("A" + tot + ":S" + tot).Font.Bold = true;
  ws.Range("A" + tot + ":S" + tot).Borders(8).LineStyle = 1;

  // formatovani
  var h = ws.Range("A1:S1");
  h.Font.Bold = true;
  h.Font.Color = rgb(255, 255, 255);
  h.Interior.Color = rgb(31, 78, 120);
  h.WrapText = true;
  h.HorizontalAlignment = -4108;
  h.VerticalAlignment = -4108;
  ws.Range("F2:Q" + tot).NumberFormat = "#,##0";
  ws.Range("R2:S" + tot).NumberFormat = "General";
  ws.Cells.Font.Name = "Calibri";
  ws.Cells.Font.Size = 10;
  ws.Range("A1:S" + last).AutoFilter();
  ws.Activate();
  ws.Range("C2").Select();
  xl.ActiveWindow.FreezePanes = true;
  ws.Range("A1").Select();

  // kontrolni list
  var ks = wb.Worksheets.Add(null, ws);
  ks.Name = "Kontrola sou\u010Dt\u016F";
  var kh = ["Polo\u017Eka", "Souhrn v XML (\u010CSSZ)", "Sou\u010Det z formul\u00E1\u0159\u016F", "Rozd\u00EDl", "Tolerance (K\u010D)", "Stav"];
  for (c = 0; c < kh.length; c++) ks.Cells(1, c + 1).Value = kh[c];

  function sv(xp) {
    var q = doc.selectSingleNode(xp);
    return q === null ? 0 : parseNum(q.text);
  }
  var items = [
    ["Z\u00E1loha na da\u0148 po slev\u00E1ch", sv("//so:danZalohaPoSleve"), "='Zam\u011Bstnanci'!J" + tot, 1],
    ["Pojistn\u00E9 zam\u011Bstnanc\u016F (SP)", sv("//pvpoj:pojistneZamestnance"), "='Zam\u011Bstnanci'!O" + tot, 1],
    ["Pojistn\u00E9 zam\u011Bstnavatele (SP) celkem", sv("//pvpoj:pojistneZamestnavateleCelkem"), "='Zam\u011Bstnanci'!P" + tot, cnt],
    ["\u00DAhrn VZ zam\u011Bstnavatele (nerizikov\u00ED)", sv("//pvpoj:zakladZamestnavateleA"), "='Zam\u011Bstnanci'!Q" + tot, 1]
  ];
  var j;
  for (j = 0; j < items.length; j++) {
    var rr = j + 2;
    ks.Cells(rr, 1).Value = items[j][0];
    ks.Cells(rr, 2).Value = items[j][1];
    ks.Cells(rr, 3).Formula = items[j][2];
    ks.Cells(rr, 4).Formula = "=C" + rr + "-B" + rr;
    ks.Cells(rr, 5).Value = items[j][3];
    ks.Cells(rr, 6).Formula = "=IF(ABS(D" + rr + ")<=E" + rr + ",\"OK\",\"ROZD\u00CDL\")";
  }
  ks.Cells(7, 1).Value = "Pozn\u00E1mka: sloupec \u201ESouhrn v XML\u201C je p\u0159evzat ze souhrnn\u00E9 \u010D\u00E1sti souboru, sou\u010Dty z formul\u00E1\u0159\u016F jsou vzorce. " +
    "Pojistn\u00E9 zam\u011Bstnavatele se u formul\u00E1\u0159\u016F zaokrouhluje nahoru zvl\u00E1\u0161\u0165, proto je tolerance rovna po\u010Dtu formul\u00E1\u0159\u016F.";
  var kh2 = ks.Range("A1:F1");
  kh2.Font.Bold = true;
  kh2.Font.Color = rgb(255, 255, 255);
  kh2.Interior.Color = rgb(31, 78, 120);
  ks.Cells.Font.Name = "Calibri";
  ks.Cells.Font.Size = 10;
  ks.Range("B2:D5").NumberFormat = "#,##0";
  // \u0161\u00ED\u0159ky sloupc\u016F: nejd\u0159\u00EDv cca 1 cm, pak automaticky (v\u010Detn\u011B zalomen\u00E9ho z\u00E1hlav\u00ED)
  ws.Range("A1:S" + tot).ColumnWidth = 4.7;
  ws.Range("A1:S" + tot).Columns.AutoFit();
  // v\u00FD\u0161ka prvn\u00EDho \u0159\u00E1dku automaticky podle zalomen\u00E9ho z\u00E1hlav\u00ED
  ws.Rows(1).AutoFit();
  ks.Range("A1:F5").Columns.AutoFit();
  ws.Activate();

  var outPath = fso.BuildPath(fso.GetParentFolderName(xmlPath), fso.GetBaseName(xmlPath) + ".xlsx");
  wb.SaveAs(outPath, 51);
  xl.Calculate();
  var stav = "";
  for (j = 0; j < items.length; j++) stav += "\n" + items[j][0] + ": " + ks.Cells(j + 2, 6).Value;
  wb.Close(false);
  xl.Quit();
  WScript.Echo("Hotovo. Zpracov\u00E1no formul\u00E1\u0159\u016F: " + cnt + "\nV\u00FDstup: " + outPath + "\n\nKontrola:" + stav);
} catch (e2) {
  try { xl.DisplayAlerts = false; xl.Quit(); } catch (e3) {}
  WScript.Echo("Chyba: " + e2.message);
  WScript.Quit(1);
}
