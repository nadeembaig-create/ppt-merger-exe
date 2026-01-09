from pptx import Presentation
from tkinter import Tk, filedialog, messagebox
import os

def merge_presentations(files):
    merged = Presentation()

    while len(merged.slides) > 0:
        rId = merged.slides._sldIdLst[0].rId
        merged.part.drop_rel(rId)
        del merged.slides._sldIdLst[0]

    for file in files:
        prs = Presentation(file)
        for slide in prs.slides:
            slide_layout = merged.slide_layouts[6]
            new_slide = merged.slides.add_slide(slide_layout)

            for shape in slide.shapes:
                if shape.has_text_frame:
                    textbox = new_slide.shapes.add_textbox(
                        shape.left, shape.top, shape.width, shape.height
                    )
                    textbox.text = shape.text

    output = os.path.join(os.path.dirname(files[0]), "Merged_Presentation.pptx")
    merged.save(output)
    messagebox.showinfo("Success", f"Saved:\n{output}")

def select_files():
    files = filedialog.askopenfilenames(
        title="Select PowerPoint Files",
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    if files:
        merge_presentations(files)

root = Tk()
root.withdraw()
select_files()
