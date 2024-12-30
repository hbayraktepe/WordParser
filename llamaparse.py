from llama_parse import LlamaParse

parser = LlamaParse(
    api_key="llx-LK1OLvFyK5Px5JRnL7Kr2Px9JPK1FRDVp4IVNfwVCe3d1wHR",  # can also be set in your env as LLAMA_CLOUD_API_KEY
    result_type = "markdown",  # "markdown" and "text" are available
    num_workers=4,  # if multiple files passed, split in `num_workers` API calls
    verbose=True,
    language="en",  # Optionally you can define a language, default=en
)

# sync
documents = parser.load_data("TestFiles/complex_test1/complex_test1.md")

print(documents)



