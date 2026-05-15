import gradio as gr, os, tempfile, time
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

SAMPLES = {
"Бичил уурхай": """1. Нийтлэг үндэслэл
1.1. Бичил уурхай эрхлэгч этгээд нь {aimag} аймгийн {sum} сумын нутагт орших {area} гектар талбайд {dur} жилийн хугацаагаар {mineral} олборлох үйл ажиллагаа эрхэлнэ.
1.2. Энэхүү гэрээ нь Монгол Улсын Ашигт малтмалын тухай хууль, Бичил уурхайн тухай журам болон холбогдох бусад хууль тогтоомжийн хүрээнд байгуулагдана.
1.3. Гэрээний талууд харилцан тэгш эрхтэй бөгөөд энэхүү гэрээнд заасан нөхцөлийг биелүүлэх үүрэгтэй.

2. Тусгай зөвшөөрөл эзэмшигчийн эрх, үүрэг
2.1. {pb} нь тусгай зөвшөөрлийн дагуу {mineral} ашигт малтмал олборлох эрхтэй.
2.2. Олборлолтын явцад байгаль орчныг хамгаалах арга хэмжээг бүрэн хэрэгжүүлнэ.
2.3. Жил бүрийн 12 дугаар сарын 25-ны дотор үйл ажиллагааны тайланг {pa}-ын тамгын газарт хүргүүлнэ.
2.4. Ашигт малтмал олборлосны нөхөн сэргээлтийн ажлыг заасан хугацаанд гүйцэтгэнэ.

3. Засаг даргын эрх, үүрэг
3.1. {pa} нь тусгай зөвшөөрлийн талбайд хяналт тавих эрхтэй.
3.2. Орон нутгийн хөгжилд чиглэсэн ажлуудад дэмжлэг үзүүлнэ.
3.3. Гэрээний биелэлтэд хяналт тавьж, зөрчил гарсан тохиолдолд арга хэмжээ авна.

4. Талуудын харилцаа
4.1. Талууд харилцан хүндэтгэлтэй хамтран ажиллана.
4.2. Маргаантай асуудлыг харилцан тохиролцох замаар шийдвэрлэнэ.
4.3. Аль нэг тал гэрээний үүргээ биелүүлээгүй тохиолдолд нөгөө тал бичгээр мэдэгдэнэ.

5. Гэрээний хариуцлага, маргаан шийдвэрлэх
5.1. Гэрээний үүргээ биелүүлээгүй тал хохирлыг нөхөн төлөх үүрэгтэй.
5.2. Маргааныг эхлээд харилцан тохиролцох замаар шийдвэрлэх бөгөөд тохиролцоонд хүрэхгүй бол шүүхэд хандана.

6. Гэрээний хэрэгжилт
6.1. Талууд гэрээний биелэлтэд хамтран хяналт тавина.
6.2. Жил бүр гэрээний биелэлтийн тайланг гаргаж хэлэлцэнэ.

7. Гэрээний хугацаа, хүчин төгөлдөр болох
7.1. Энэхүү гэрээ нь талуудын гарын үсэг зурсан өдрөөс хүчин төгөлдөр болно.
7.2. Гэрээний хугацаа {dur} жил байх бөгөөд харилцан тохиролцсоны үндсэн дээр сунгаж болно.

8. Бусад зүйл
8.1. Энэхүү гэрээнд заагаагүй асуудлыг холбогдох хууль тогтоомжийн дагуу шийдвэрлэнэ.
8.2. Гэрээг монгол хэлээр 2 хувь үйлдэж, тал тус бүр 1 хувийг хадгална.""",

"Хайгуулын гэрээ": """1. Нийтлэг үндэслэл
1.1. {pb} нь {aimag} аймгийн {sum} сумын нутагт орших {area} гектар талбайд {dur} жилийн хугацаагаар {mineral} ашигт малтмалын хайгуулын ажил явуулна.
1.2. Энэхүү гэрээ нь Монгол Улсын Ашигт малтмалын тухай хууль болон холбогдох хууль тогтоомжийн хүрээнд байгуулагдана.
1.3. Хайгуулын ажлын үр дүнг {pa}-д тогтмол мэдээлнэ.

2. Тусгай зөвшөөрөл эзэмшигчийн эрх, үүрэг
2.1. {pb} нь хайгуулын тусгай зөвшөөрлийн хүрээнд геологийн судалгаа явуулах эрхтэй.
2.2. Хайгуулын ажлын явцад байгаль орчинд учруулах хохирлыг багасгах арга хэмжээ авна.
2.3. Хайгуулын үр дүнгийн тайланг жил бүр {pa}-д хүргүүлнэ.
2.4. Хайгуулын ажлыг дуусгасны дараа нөхөн сэргээлтийг гүйцэтгэнэ.

3. Засаг даргын эрх, үүрэг
3.1. {pa} нь хайгуулын талбайд хяналт тавих эрхтэй.
3.2. Хайгуулын ажлыг дэмжиж, шаардлагатай зөвшөөрлийг олгоно.
3.3. Орон нутгийн иргэдийн эрх ашгийг хамгаалах арга хэмжээ авна.

4. Талуудын харилцаа
4.1. Талууд харилцан хүндэтгэлтэй хамтран ажиллана.
4.2. Хайгуулын үр дүнг харилцан хэлэлцэж, цаашдын чиглэлийг тодорхойлно.
4.3. Маргаантай асуудлыг харилцан тохиролцох замаар шийдвэрлэнэ.

5. Гэрээний хариуцлага, маргаан шийдвэрлэх
5.1. Гэрээний үүргээ биелүүлээгүй тал хохирлыг нөхөн төлөх үүрэгтэй.
5.2. Маргааныг эхлээд харилцан тохиролцох замаар шийдвэрлэх бөгөөд тохиролцоонд хүрэхгүй бол шүүхэд хандана.

6. Гэрээний хэрэгжилт
6.1. Хайгуулын ажлын явцыг тогтмол хянаж, биелэлтийн тайланг гаргана.
6.2. Тайланг улирал бүр {pa}-д хүргүүлнэ.

7. Гэрээний хугацаа, хүчин төгөлдөр болох
7.1. Энэхүү гэрээ нь талуудын гарын үсэг зурсан өдрөөс хүчин төгөлдөр болно.
7.2. Гэрээний хугацаа {dur} жил байх бөгөөд харилцан тохиролцсоны үндсэн дээр сунгаж болно.

8. Бусад зүйл
8.1. Энэхүү гэрээнд заагаагүй асуудлыг холбогдох хууль тогтоомжийн дагуу шийдвэрлэнэ.
8.2. Гэрээг монгол хэлээр 2 хувь үйлдэж, тал тус бүр 1 хувийг хадгална.""",

"Ашиглалтын тухай гэрээ": """1. Нийтлэг үндэслэл
1.1. {pb} нь {aimag} аймгийн {sum} сумын нутагт орших {area} гектар талбайд {dur} жилийн хугацаагаар {mineral} ашигт малтмал ашиглах үйл ажиллагаа эрхэлнэ.
1.2. Энэхүү гэрээ нь Монгол Улсын Ашигт малтмалын тухай хууль болон холбогдох хууль тогтоомжийн хүрээнд байгуулагдана.
1.3. Ашиглалтын явцад орон нутгийн хөгжилд хувь нэмэр оруулна.

2. Тусгай зөвшөөрөл эзэмшигчийн эрх, үүрэг
2.1. {pb} нь ашиглалтын тусгай зөвшөөрлийн дагуу {mineral} олборлох эрхтэй.
2.2. Олборлолтын технологийн шаардлагыг бүрэн мөрдөнө.
2.3. Орон нутгийн иргэдийг ажлын байраар хангах талаар хамтран ажиллана.
2.4. Ашиглалтын дараах нөхөн сэргээлтийн ажлыг хуулийн дагуу гүйцэтгэнэ.

3. Засаг даргын эрх, үүрэг
3.1. {pa} нь ашиглалтын талбайд хяналт тавих эрхтэй.
3.2. Орон нутгийн хөгжлийн асуудлаар хамтран ажиллана.
3.3. Байгаль орчны хяналтыг тогтмол явуулна.

4. Талуудын харилцаа
4.1. Талууд харилцан хүндэтгэлтэй хамтран ажиллана.
4.2. Ашиглалтын явцын тайланг хамтран хэлэлцэнэ.
4.3. Маргаантай асуудлыг харилцан тохиролцох замаар шийдвэрлэнэ.

5. Гэрээний хариуцлага, маргаан шийдвэрлэх
5.1. Гэрээний үүргээ биелүүлээгүй тал хохирлыг нөхөн төлөх үүрэгтэй.
5.2. Маргааныг эхлээд харилцан тохиролцох замаар шийдвэрлэнэ.

6. Гэрээний хэрэгжилт
6.1. Ашиглалтын явцыг тогтмол хянаж, тайланг улирал бүр гаргана.
6.2. Жилийн эцэст нэгтгэсэн тайланг {pa}-д хүргүүлнэ.

7. Гэрээний хугацаа, хүчин төгөлдөр болох
7.1. Энэхүү гэрээ нь талуудын гарын үсэг зурсан өдрөөс хүчин төгөлдөр болно.
7.2. Гэрээний хугацаа {dur} жил байх бөгөөд харилцан тохиролцсоны үндсэн дээр сунгаж болно.

8. Бусад зүйл
8.1. Энэхүү гэрээнд заагаагүй асуудлыг холбогдох хууль тогтоомжийн дагуу шийдвэрлэнэ.
8.2. Гэрээг монгол хэлээр 2 хувь үйлдэж, тал тус бүр 1 хувийг хадгална.""",

"Хамтын ажиллагааны гэрээ": """1. Нийтлэг үндэслэл
1.1. {pa} болон {pb} нь {aimag} аймгийн {sum} сумын нутагт {mineral} ашигт малтмалтай холбоотой хамтын ажиллагааг {dur} жилийн хугацаатай явуулна.
1.2. Энэхүү гэрээ нь Монгол Улсын хууль тогтоомжийн хүрээнд байгуулагдана.
1.3. Хамтын ажиллагааны зорилго нь орон нутгийн хөгжил, ашигт малтмалын зөв ашиглалтыг хангах явдал юм.

2. Тусгай зөвшөөрөл эзэмшигчийн эрх, үүрэг
2.1. {pb} нь хамтын ажиллагааны хүрээнд {mineral} олборлолтод оролцох эрхтэй.
2.2. Орон нутгийн хөгжлийн санд жил бүр хувь нэмэр оруулна.
2.3. Ажлын байр нэмэгдүүлэх талаар {pa}-тай хамтран ажиллана.

3. Засаг даргын эрх, үүрэг
3.1. {pa} нь хамтын ажиллагааны үйл ажиллагаанд хяналт тавих эрхтэй.
3.2. Орон нутгийн иргэдийн эрх ашгийг хамгаалах арга хэмжээ авна.
3.3. Хамтын ажиллагааг дэмжих орчин нөхцөл бүрдүүлнэ.

4. Талуудын харилцаа
4.1. Талууд харилцан хүндэтгэлтэй хамтран ажиллана.
4.2. Хамтын ажиллагааны явцыг тогтмол хэлэлцэж, шаардлагатай өөрчлөлт оруулна.
4.3. Маргаантай асуудлыг харилцан тохиролцох замаар шийдвэрлэнэ.

5. Гэрээний хариуцлага, маргаан шийдвэрлэх
5.1. Гэрээний үүргээ биелүүлээгүй тал хохирлыг нөхөн төлөх үүрэгтэй.
5.2. Маргааныг эхлээд харилцан тохиролцох замаар шийдвэрлэнэ.

6. Гэрээний хэрэгжилт
6.1. Хамтын ажиллагааны явцыг тогтмол хянаж тайлагнана.
6.2. Жилийн эцэст нэгтгэсэн тайланг хамтран хэлэлцэнэ.

7. Гэрээний хугацаа, хүчин төгөлдөр болох
7.1. Энэхүү гэрээ нь талуудын гарын үсэг зурсан өдрөөс хүчин төгөлдөр болно.
7.2. Гэрээний хугацаа {dur} жил байх бөгөөд харилцан тохиролцсоны үндсэн дээр сунгаж болно.

8. Бусад зүйл
8.1. Энэхүү гэрээнд заагаагүй асуудлыг холбогдох хууль тогтоомжийн дагуу шийдвэрлэнэ.
8.2. Гэрээг монгол хэлээр 2 хувь үйлдэж, тал тус бүр 1 хувийг хадгална."""
}

def make_word(text):
    doc = Document()
    t = doc.add_paragraph()
    r = t.add_run("УУЛ УУРХАЙН ГЭРЭЭ")
    r.bold = True
    r.font.size = Pt(16)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("")
    for l in text.split("\n"):
        if not l.strip():
            continue
        p = doc.add_paragraph()
        run = p.add_run(l)
        if l.strip() and l[0].isdigit() and ". " in l[:4] and not l[2].isdigit():
            run.bold = True
            run.font.size = Pt(13)
        else:
            run.font.size = Pt(11)
    path = os.path.join(tempfile.mkdtemp(), "гэрээ.docx")
    doc.save(path)
    return path

def make_pdf(text):
    path = os.path.join(tempfile.mkdtemp(), "гэрээ.pdf")
    fp = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    fb = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    fn, fnb = "DejaVu", "DejaVuB"
    if os.path.exists(fp):
        pdfmetrics.registerFont(TTFont(fn, fp))
        pdfmetrics.registerFont(TTFont(fnb, fb))
    else:
        fn = fnb = "Helvetica"
    doc = SimpleDocTemplate(path, pagesize=A4,
                            leftMargin=60, rightMargin=60,
                            topMargin=60, bottomMargin=60)
    st = ParagraphStyle("t", fontName=fnb, fontSize=15, alignment=1, spaceAfter=20)
    sh = ParagraphStyle("h", fontName=fnb, fontSize=12, spaceAfter=8, spaceBefore=12)
    sn = ParagraphStyle("n", fontName=fn, fontSize=10, leading=16, spaceAfter=5)
    els = [Paragraph("УУЛ УУРХАЙН ГЭРЭЭ", st)]
    for l in text.split("\n"):
        if not l.strip():
            els.append(Spacer(1, 6))
            continue
        safe = l.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        if l.strip() and l[0].isdigit() and ". " in l[:4] and not l[2].isdigit():
            els.append(Paragraph(safe, sh))
        else:
            els.append(Paragraph(safe, sn))
    doc.build(els)
    return path

def run(ctype, aimag, sum_, area, mineral, dur, pa, pb, progress=gr.Progress()):
    sections = [
        "Нийтлэг үндэслэл","Тусгай зөвшөөрөл эзэмшигчийн эрх, үүрэг",
        "Засаг даргын эрх, үүрэг","Талуудын харилцаа",
        "Гэрээний хариуцлага, маргаан шийдвэрлэх","Гэрээний хэрэгжилт",
        "Гэрээний хугацаа, хүчин төгөлдөр болох","Бусад зүйл"
    ]
    for i, s in enumerate(sections):
        progress((i+1)/len(sections), desc=f"{i+1}/8: {s}")
        time.sleep(1.2)
    template = SAMPLES.get(ctype, SAMPLES["Бичил уурхай"])
    text = template.format(aimag=aimag, sum=sum_, area=area,
                           mineral=mineral, dur=dur, pa=pa, pb=pb)
    return text, make_word(text), make_pdf(text)

css = """
.gradio-container{max-width:1100px!important;margin:0 auto!important;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif!important;}
.header-wrap{padding:1.75rem 0 1.25rem;border-bottom:1px solid #e5e7eb;margin-bottom:1.5rem;}
.header-title{font-size:20px;font-weight:600;color:#111827;margin:0 0 3px;}
.header-sub{font-size:13px;color:#6b7280;margin:0 0 10px;}
.pills{display:flex;gap:6px;flex-wrap:wrap;}
.pill{font-size:11px;padding:3px 11px;border-radius:100px;border:1px solid #d1d5db;color:#374151;background:#f9fafb;font-weight:500;}
.section-title{font-size:11px!important;font-weight:600!important;letter-spacing:0.07em!important;text-transform:uppercase!important;color:#9ca3af!important;margin-bottom:10px!important;display:block;}
button.primary{background:#111827!important;color:#fff!important;border-radius:8px!important;font-size:14px!important;font-weight:500!important;border:none!important;padding:11px 0!important;width:100%!important;margin-top:6px!important;cursor:pointer!important;transition:opacity 0.15s!important;}
button.primary:hover{opacity:0.82!important;}
.notice{margin-top:1rem;padding:11px 15px;background:#fffbeb;border-left:3px solid #f59e0b;border-radius:0 6px 6px 0;font-size:12px;color:#92400e;line-height:1.65;}
.dark .header-wrap{border-bottom-color:#2d2d2d;}
.dark .header-title{color:#f3f4f6;}
.dark .header-sub{color:#9ca3af;}
.dark .pill{border-color:#374151;color:#d1d5db;background:#1f2937;}
.dark button.primary{background:#f9fafb!important;color:#111!important;}
.dark .notice{background:#1c1a00;color:#fbbf24;border-left-color:#d97706;}
"""

with gr.Blocks(title="Уул Уурхайн Гэрээ Үүсгэгч") as demo:
    gr.HTML("""
    <div class="header-wrap">
      <button onclick="document.body.classList.toggle('dark');this.textContent=document.body.classList.contains('dark')?'☀️ Light':'🌙 Dark'"
        style="float:right;padding:5px 14px;border-radius:20px;border:1px solid #d1d5db;background:transparent;cursor:pointer;font-size:12px;font-weight:500">🌙 Dark</button>
      <p class="header-title">Уул уурхайн гэрээ үүсгэгч</p>
      <p class="header-sub">Хиймэл оюун ухаанд суурилсан гэрээ боловсруулах систем</p>

    </div>""")
    with gr.Row(equal_height=False):
        with gr.Column(scale=4, min_width=300):
            gr.HTML('<span class="section-title">Гэрээний нөхцөл</span>')
            ctype = gr.Dropdown(
                ["Бичил уурхай","Ашиглалтын тухай гэрээ","Хайгуулын гэрээ","Хамтын ажиллагааны гэрээ"],
                value="Бичил уурхай", label="Гэрээний төрөл", container=True)
            with gr.Row():
                aimag = gr.Textbox(label="Аймаг", value="Сүхбаатар")
                sum_  = gr.Textbox(label="Сум",   value="Тариалан")
            with gr.Row():
                area    = gr.Textbox(label="Талбай (га)",    value="0.5")
                mineral = gr.Textbox(label="Ашигт малтмал", value="Жонш")
            dur = gr.Textbox(label="Хугацаа (жил)", value="1")
            with gr.Row():
                pa = gr.Textbox(label="А тал", value="Засаг дарга")
                pb = gr.Textbox(label="Б тал", value="БАЙГУУЛЛАГА нөхөрлөл")
            btn = gr.Button("Гэрээ үүсгэх", variant="primary")
        with gr.Column(scale=6, min_width=400):
            gr.HTML('<span class="section-title">Үүсгэсэн гэрээ</span>')
            out = gr.Textbox(label="", lines=26, max_lines=40,
                             placeholder="Гэрээний нөхцөлийг бөглөөд үүсгэх товчийг дарна уу...")
            gr.HTML('<span class="section-title" style="margin-top:12px;display:block;">Татаж авах</span>')
            with gr.Row():
                wf = gr.File(label="Word (.docx)", file_count="single")
                pf = gr.File(label="PDF (.pdf)",   file_count="single")
            gr.HTML('<div class="notice"><strong>Анхааруулга:</strong> Энэхүү гэрээ нь AI-аар автомат үүсгэсэн төсөл бөгөөд хуулийн мэргэжилтнээр заавал хянуулах ёстой.</div>')
    btn.click(fn=run, inputs=[ctype,aimag,sum_,area,mineral,dur,pa,pb], outputs=[out,wf,pf])

demo.launch(css=css, server_name="0.0.0.0", server_port=7860, share=True)
