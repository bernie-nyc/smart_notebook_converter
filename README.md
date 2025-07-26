# SMART Notebook to PowerPoint Converter

This repository contains a Python script that converts **SMART Notebook (`.notebook`) files** into **Microsoft PowerPoint (`.pptx`) presentations**. Each Notebook file is a ZIP archive that holds individual page assets in **SVG**, **PNG** or **JPEG** formats. The script scans a directory tree for `.notebook` files, extracts all supported page assets, converts SVG pages to PNG if necessary, and assembles a static slide deck using [`python‑pptx`](https://python-pptx.readthedocs.io/). Interactive SMART features (animations, audio, video, embedded quizzes) are not exported.

## Features

* **Batch processing:** provide a root directory and the script recursively converts every `.notebook` file it finds. Use `--output-dir` to send all results to a single directory; omit it to write each `.pptx` alongside its source file.
* **Asset handling:** pages stored as **SVG** are rasterised to PNG; existing PNG/JPEG pages are used directly. Unsupported page formats (proprietary XML, interactive objects) are skipped.
* **Fallback conversion:** if SVG rasterisation fails (e.g. due to missing Cairo), the script attempts to call ImageMagick’s `magick` command instead.

## Usage

Install the Python dependencies (see below) and run:

```bash
python notebook_to_ppt.py --input <path> [--output-dir <output_directory>] [--verbose]
```

* `<path>` can be a single `.notebook` file or a directory; the script recurses through subdirectories.
* If `--output-dir` is not supplied, each PowerPoint file is saved in the same folder as its source `.notebook` file.
* Use `--verbose` for detailed logging.

## Dependencies

The converter script depends on:

1. **Python 3** with the [python‑pptx](https://python-pptx.readthedocs.io/) library.

2. **CairoSVG** to rasterise SVG pages. CairoSVG is a Python package that relies on the C‑level Cairo library and FFI headers. The CairoSVG documentation notes that additional tools are required during installation—specifically Cairo and FFI headers—and the package names vary by operating system:

   * **Windows:** install Cairo (for example via GTK) and the Microsoft Visual C++ compiler. An easier option is to install the prebuilt wheel that bundles Cairo: `pip install "cairosvg==2.5.2"`.
   * **macOS:** use Homebrew to install `cairo` and `libffi`, then run `pip install cairosvg`.
   * **Linux (Debian/Ubuntu):** install the `cairo`, `python3-dev` and `libffi-dev` packages, then run `pip install cairosvg`.

   On Windows you must also ensure that the `libcairo-2.dll` file is on your **PATH**. The cairocffi documentation notes that Cairo must be available as a shared library and suggests using **Alexander Shaduri’s GTK+ installer**, which places `libcairo-2.dll` on your PATH.

3. **ImageMagick** (fallback). If CairoSVG cannot rasterise a page, the script falls back to ImageMagick’s `magick` command. You need a version of ImageMagick that supports SVG:

   * **Windows:** the official site offers a self‑installing Q16 HDRI build (16 bits per component). According to ImageMagick’s download page, the Windows version is self‑installing—just click the appropriate file (e.g. `ImageMagick‑7.x.x‑Q16‑HDRI‑x64‑dll.exe`) and follow the prompts.
   * **macOS:** install via Homebrew using `brew install imagemagick`; this downloads prebuilt binaries and their delegate libraries.
   * **Linux (Debian/Ubuntu):** install via your package manager, e.g. `sudo apt‑get install imagemagick`.

   After installation, verify that the `magick` command is on your PATH and can convert an SVG: run `magick input.svg output.png`. If this fails, your ImageMagick build lacks SVG support.

## Troubleshooting Cairo errors

CairoSVG relies on the C‑level **Cairo** graphics library. If you receive errors such as:

```
no library called "cairo-2" was found
no library called "libcairo.so.2" was found
no library called "libcairo-2.dll" was found
```

the Cairo DLL or shared library cannot be located. To resolve this:

* **Windows:** ensure that `libcairo-2.dll` exists on your system and that its directory is included in the `PATH`. The cairocffi documentation recommends using Alexander Shaduri’s GTK+ installer and keeping the “set up PATH” checkbox checked. Alternatively, install the CairoSVG wheel `cairosvg==2.5.2`, which bundles the necessary DLLs.

* **macOS:** install `cairo` and `libffi` via Homebrew (`brew install cairo libffi`). On Apple Silicon systems, Homebrew installs libraries under `/opt/homebrew/lib`; create a symlink from `libcairo.2.dylib` to `/usr/local/lib` or update your `DYLD_LIBRARY_PATH` accordingly. For example:

  ```bash
  ln -s /opt/homebrew/lib/libcairo.2.dylib /usr/local/lib/libcairo.2.dylib
  ```

* **Linux:** install the `libcairo2`, `libcairo2-dev` (or `python3-dev`, `libffi-dev`) packages via your package manager and ensure `LD_LIBRARY_PATH` includes the directory containing `libcairo.so.2`.

Once the Cairo library can be located by your system loader, the warnings disappear and SVG pages will convert correctly.

## Repository contents

* **`notebook_to_ppt.py`** – the Python script that performs the conversion.
* **`ImageMagick‑7.x.x‑Q16‑HDRI‑x64‑dll.exe`** (Windows only) – the official ImageMagick installer from the project’s download page. Run this executable to install ImageMagick with high dynamic‑range imaging support. During installation, ensure that the option to add ImageMagick to your system PATH is selected.

## License

This converter script is provided for educational purposes without any warranty. SMART Notebook and its associated trademarks are owned by SMART Technologies ULC. Use at your own risk.
