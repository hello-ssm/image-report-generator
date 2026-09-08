from smolagents import tool
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import os
import datetime
import gradio as gr


@tool
def generate_document_tool(image_paths: list[str], descriptions: list[str], output_format: str) -> str:
    """Generates a PDF or Excel document combining multiple images, each with its own text description.
    Args:
        image_paths: local file paths to the images to embed in the document, in order
        descriptions: text description for each image, matching image_paths by position
        output_format: desired output format, either 'pdf' or 'excel'
    """
    if len(image_paths) != len(descriptions):
        raise ValueError("image_paths and descriptions must have the same length")

    output_format = output_format.strip().lower()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    if output_format == "pdf":
        output_path = f"report_{timestamp}.pdf"
        pdf = FPDF()
        pdf.set_font("Helvetica", size=12)
        for image_path, description in zip(image_paths, descriptions):
            pdf.add_page()
            pdf.multi_cell(0, 10, description)
            pdf.ln(5)
            pdf.image(image_path, w=150)
        pdf.output(output_path)
        return output_path

    elif output_format in ("excel", "xlsx"):
        output_path = f"report_{timestamp}.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.column_dimensions["A"].width = 60
        row = 1
        for image_path, description in zip(image_paths, descriptions):
            ws.cell(row=row, column=1, value=description)
            img = XLImage(image_path)
            ws.add_image(img, f"A{row + 1}")
            row += 18  # leave room for the embedded image before the next entry
        wb.save(output_path)
        return output_path

    else:
        raise ValueError("output_format must be 'pdf' or 'excel'")


def _build_report(image_paths, descriptions, output_format):
    if not image_paths:
        raise gr.Error("请至少上传一张图片")
    if any(not d or not d.strip() for d in descriptions):
        raise gr.Error("请给每一张图片都填写描述")
    today = datetime.date.today().isoformat()
    dated_descriptions = [f"{d.strip()}\n\n生成日期: {today}" for d in descriptions]
    return generate_document_tool(image_paths=image_paths, descriptions=dated_descriptions, output_format=output_format)


with gr.Blocks(title="图文报告生成器") as demo:
    gr.Markdown(
        "## 图文报告生成器\n"
        "上传多张图片，为每张图片分别填写文字描述，一键生成可下载的 PDF / Excel 报告。\n"
        "手机浏览器打开这个页面同样可以直接拍照或从相册选图上传。"
    )
    files_input = gr.File(label="上传图片（可多选）", file_count="multiple", file_types=["image"])

    @gr.render(inputs=files_input)
    def render_form(files):
        if not files:
            gr.Markdown("请先上传图片。")
            return

        description_boxes = [
            gr.Textbox(label=f"描述 - {os.path.basename(f)}", placeholder="请输入这张图片的文字说明...")
            for f in files
        ]
        format_input = gr.Radio(choices=["pdf", "excel"], value="pdf", label="输出格式")
        generate_btn = gr.Button("生成报告", variant="primary")
        output_file = gr.File(label="生成的报告（点击下载）")

        def on_generate(*values):
            *descriptions, output_format = values
            return _build_report(files, list(descriptions), output_format)

        generate_btn.click(
            fn=on_generate,
            inputs=description_boxes + [format_input],
            outputs=output_file,
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
