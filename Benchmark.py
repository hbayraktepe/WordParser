import difflib


def compare_files(file1_path, file2_path):
    with open(file1_path, "r", encoding="utf-8") as file1, open(
        file2_path, "r", encoding="utf-8"
    ) as file2:
        file1_lines = file1.readlines()
        file2_lines = file2.readlines()

    diff = difflib.unified_diff(
        file1_lines, file2_lines, fromfile=file1_path, tofile=file2_path
    )
    return "".join(diff)


if __name__ == "__main__":
    my_markdown_path = "TestFiles/Test4/Test4.md"  # Your generated Markdown file
    pandoc_markdown_path = (
        "PandocFiles/Test4/Test4_pandoc.md"  # Pandoc generated Markdown file
    )

    diff_result = compare_files(my_markdown_path, pandoc_markdown_path)

    if diff_result:
        print("Differences found between the files:")
        print(diff_result)
    else:
        print("No differences found between the files.")


##  pandoc -s TestFiles/Test1.docx -t markdown -o TestFiles/Test1_pandoc.md --extract-media=TestFiles/Test1
