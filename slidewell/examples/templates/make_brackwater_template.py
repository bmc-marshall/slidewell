"""Builds brackwater-template.pptx: a fictional corporate PowerPoint template, made the way a real company's is
(theme colours and fonts, a slide master with logo and footer, and named layouts with placeholders). It exists so
the template example (build_brackwater_update.py) can show Slidewell working inside someone else's template.

Brackwater Group is invented. Run: python make_brackwater_template.py
"""
import copy
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Inches

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A, P = NS["a"], NS["p"]
NSDECL = f'xmlns:a="{A}" xmlns:p="{P}" xmlns:r="{NS["r"]}"'

W, H = 13.333, 7.5
M = 0.6                      # outer margin
TEAL, TEAL_DK, COPPER, INK, GREY, RULE = "0E4D5C", "0A3843", "C0703A", "1F2A30", "6B767C", "D5DADC"
COLORS = dict(dk1=INK, lt1="FFFFFF", dk2=TEAL, lt2="EEF2F3", accent1=TEAL, accent2=COPPER, accent3="5E8C61",
              accent4="D9A441", accent5="8A9BA3", accent6="B5483B", hlink=TEAL, folHlink="5E6E75")
HEAD_FONT, BODY_FONT = "Cambria", "Calibri"


def emu(i):
    return str(int(Inches(i)))


def xfrm(x, y, w, h):
    return f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'


def ph(id_, name, type_, idx, x, y, w, h, body="", anchor="t", lst=""):
    t = f' type="{type_}"' if type_ else ""
    i = f' idx="{idx}"' if idx is not None else ""
    return f'''<p:sp {NSDECL}><p:nvSpPr><p:cNvPr id="{id_}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
<p:nvPr><p:ph{t}{i}/></p:nvPr></p:nvSpPr><p:spPr>{xfrm(x, y, w, h)}</p:spPr>
<p:txBody><a:bodyPr anchor="{anchor}" lIns="0" rIns="0" tIns="0" bIns="0"><a:normAutofit/></a:bodyPr><a:lstStyle>{lst}</a:lstStyle>
<a:p>{body}<a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>'''


def rect(id_, name, x, y, w, h, fill, line=None):
    ln = f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
    return f'''<p:sp {NSDECL}><p:nvSpPr><p:cNvPr id="{id_}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>
<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>'''


def wordmark(id_, x, y, color):
    """The (invented) Brackwater logo: a copper square and a letter-spaced wordmark."""
    return rect(id_, "Logo mark", x, y + 0.035, 0.14, 0.14, COPPER) + f'''<p:sp {NSDECL}><p:nvSpPr><p:cNvPr id="{id_ + 1}" name="Logo wordmark"/>
<p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr>{xfrm(x + 0.22, y, 2.2, 0.22)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr" wrap="none"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-GB" sz="1100" b="1" spc="300">
<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{BODY_FONT}"/></a:rPr><a:t>BRACKWATER</a:t></a:r></a:p></p:txBody></p:sp>'''


def slide_num(id_, color=GREY):
    lst = (f'<a:lvl1pPr algn="r"><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
           f'<a:latin typeface="{BODY_FONT}"/></a:defRPr></a:lvl1pPr>')
    body = '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-GB"/><a:t>‹#›</a:t></a:fld>'
    return ph(id_, "Slide Number Placeholder", "sldNum", 12, W - M - 0.8, 6.98, 0.8, 0.22, body, "ctr", lst)


def set_tree(part_el, shapes_xml):
    tree = part_el.find(".//p:cSld/p:spTree", NS)
    for ch in list(tree)[2:]:            # keep nvGrpSpPr and grpSpPr
        tree.remove(ch)
    for x in shapes_xml:
        for el in etree.fromstring(f"<wrap {NSDECL}>{x}</wrap>"):   # a string may hold several shapes
            tree.append(el)


def set_bg(part_el, color):
    cSld = part_el.find("p:cSld", NS)
    old = cSld.find("p:bg", NS)
    if old is not None:
        cSld.remove(old)
    bg = etree.fromstring(f'<p:bg {NSDECL}><p:bgPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>')
    cSld.insert(0, bg)


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    master = prs.slide_master
    m = master._element

    # ---- theme: colours and fonts
    theme_part = [r.target_part for r in master.part.rels.values() if r.reltype.endswith("/theme")][0]
    th = etree.fromstring(theme_part.blob)
    th.find(".//a:themeElements", NS).getparent().set("name", "Brackwater")
    cs = th.find(".//a:clrScheme", NS); cs.set("name", "Brackwater")
    for k, v in COLORS.items():
        el = cs.find(f"a:{k}", NS)
        for c in list(el):
            el.remove(c)
        etree.SubElement(el, f"{{{A}}}srgbClr").set("val", v)
    fs = th.find(".//a:fontScheme", NS); fs.set("name", "Brackwater")
    fs.find("a:majorFont/a:latin", NS).set("typeface", HEAD_FONT)
    fs.find("a:minorFont/a:latin", NS).set("typeface", BODY_FONT)
    theme_part._blob = etree.tostring(th, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ---- master: title and body styles
    ts = m.find(".//p:txStyles/p:titleStyle/a:lvl1pPr", NS)
    for c in list(ts):
        ts.remove(c)
    ts.set("algn", "l")
    ts.append(etree.fromstring(f'<a:lnSpc {NSDECL}><a:spcPct val="95000"/></a:lnSpc>'))
    ts.append(etree.fromstring(f'<a:defRPr {NSDECL} sz="2800" b="0"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill>'
                               f'<a:latin typeface="+mj-lt"/></a:defRPr>'))
    bs = m.find(".//p:txStyles/p:bodyStyle", NS)
    sizes = {1: 1800, 2: 1600, 3: 1400, 4: 1400, 5: 1400}
    for lvl, sz in sizes.items():
        el = bs.find(f"a:lvl{lvl}pPr", NS)
        for c in list(el):
            el.remove(c)
        ind = 0.25 * lvl
        el.set("marL", emu(ind)); el.set("indent", emu(-0.25))
        el.append(etree.fromstring(f'<a:spcBef {NSDECL}><a:spcPts val="{600 if lvl == 1 else 300}"/></a:spcBef>'))
        el.append(etree.fromstring(f'<a:buClr {NSDECL}><a:schemeClr val="accent2"/></a:buClr>'))
        el.append(etree.fromstring(f'<a:buFont {NSDECL} typeface="Arial"/>'))
        el.append(etree.fromstring(f'<a:buChar {NSDECL} char="{"•" if lvl % 2 else "–"}"/>'))
        el.append(etree.fromstring(f'<a:defRPr {NSDECL} sz="{sz}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
                                   f'<a:latin typeface="+mn-lt"/></a:defRPr>'))
    set_bg(m, "FFFFFF")
    set_tree(m, [
        rect(20, "Title accent", M, 0.42, 0.55, 0.06, COPPER),
        ph(2, "Title Placeholder", "title", None, M, 0.62, W - 2 * M, 0.95, anchor="t"),
        ph(3, "Text Placeholder", "body", 1, M, 1.95, W - 2 * M, 4.6),
        rect(21, "Footer rule", M, 6.85, W - 2 * M, 0.012, RULE),
        wordmark(22, M, 6.98, TEAL),
        slide_num(5),
    ])

    # ---- layouts: keep five, rebuild each
    keep = {"Title Slide": "Cover", "Title and Content": "Title and content", "Section Header": "Section divider",
            "Title Only": "Title only", "Blank": "Blank"}
    layouts = master.slide_layouts
    for lay in list(layouts):
        if lay.name not in keep:
            layouts.remove(lay)
    for lay in layouts:
        el = lay._element
        new = keep[lay.name]
        el.find("p:cSld", NS).set("name", new)
        if new == "Cover":
            set_bg(el, TEAL)
            lst_t = (f'<a:lvl1pPr><a:defRPr sz="4000"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:lvl1pPr>')
            lst_s = (f'<a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="D6E4E7"/></a:solidFill>'
                     f'</a:defRPr></a:lvl1pPr>')
            set_tree(el, [
                rect(10, "Cover accent", M, 2.35, 0.9, 0.08, COPPER),
                ph(2, "Title 1", "ctrTitle", None, M, 2.65, 8.4, 1.6, anchor="t", lst=lst_t),
                ph(3, "Subtitle 2", "subTitle", 1, M, 4.45, 8.4, 0.9, lst=lst_s),
                wordmark(11, M, 6.78, "FFFFFF"),
                rect(13, "Cover panel", 9.6, 0, W - 9.6, H, TEAL_DK),
            ])
            el.set("showMasterSp", "0")
        elif new == "Section divider":
            set_bg(el, "FFFFFF")
            lst_t = '<a:lvl1pPr><a:defRPr sz="3600"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:lvl1pPr>'
            lst_b = ('<a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="D6E4E7"/></a:solidFill>'
                     '</a:defRPr></a:lvl1pPr>')
            set_tree(el, [
                rect(10, "Section panel", 0, 0, W, H, TEAL),
                rect(14, "Section panel shade", 9.6, 0, W - 9.6, H, TEAL_DK),
                rect(11, "Section accent", M, 2.75, 0.9, 0.08, COPPER),
                ph(2, "Title 1", "title", None, M, 3.05, 8.4, 1.4, anchor="t", lst=lst_t),
                ph(3, "Text 2", "body", 1, M, 4.6, 8.4, 0.9, lst=lst_b),
                wordmark(12, M, 6.98, "FFFFFF"),
                slide_num(5),
            ])
            el.set("showMasterSp", "0")
        elif new == "Title and content":
            set_tree(el, [ph(2, "Title 1", "title", None, M, 0.62, W - 2 * M, 0.95),
                          ph(3, "Content Placeholder 2", None, 1, M, 1.95, W - 2 * M, 4.6),
                          slide_num(5)])
        elif new == "Title only":
            set_tree(el, [ph(2, "Title 1", "title", None, M, 0.62, W - 2 * M, 0.95), slide_num(5)])
        else:
            set_tree(el, [slide_num(5)])

    cp = prs.core_properties
    cp.author = cp.last_modified_by = "Brackwater Group (fictional)"
    cp.title, cp.revision = "Brackwater template", 1
    out = Path(__file__).with_name("brackwater-template.pptx")
    prs.save(out)
    print(out)


if __name__ == "__main__":
    main()
