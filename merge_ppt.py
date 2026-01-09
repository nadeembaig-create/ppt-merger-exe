import os
from pptx import Presentation
from tkinter import Tk, filedialog, simpledialog, messagebox
from copy import deepcopy

def clone_slide(prs, slide):
    layout = prs.slide_layouts[slide.slide_layout.slide_layout_id]
    new_slide = prs.slides.add_slide(layout)

    for shape in slide.shapes:
        el = shape.element
        new_el = deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(
            new_el, 'p:extLst'
        )

def merge_presentations(files, output_file):
    merged = Presentation()

    # Remove default slide
    while merged.slides:
        rId = merged.slides._sldIdLst[0].rId
        merged.part.drop_rel(rId)
        del merged.slides._sldIdLst[0]

    for file in files:
        prs = Presentation(file)
        for slide in prs.slides:
            clone_slide(merged, slide)

    merged.save(output_file)
    messagebox.showinfo("Success", f"Merged file created:\n{output_file}")

def select_files():
    files = filedialog.askopenfilenames(
        title="Select PowerPoint files",
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    return list(files)

def select_folder():
    folder = filedialog.askdirectory(title="Select folder containing PPT files")
    if not folder:
        return []
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".pptx")
    ]

def main():
    root = Tk()
    root.withdraw()

    choice = messagebox.askyesno(
        "Merge Method",
        "YES = Merge entire folder\nNO = Select individual files"
    )

    files = select_folder() if choice else select_files()

    if not files:
        messagebox.showerror("Error", "No PowerPoint files selected.")
        return

    name = simpledialog.askstring(
        "Output File",
        "Enter output filename (without extension):"
    )

    if not name:
        return

    output = os.path.join(os.path.dirname(files[0]), name + ".pptx")
    merge_presentations(files, output)

if __name__ == "__main__":
    main()
