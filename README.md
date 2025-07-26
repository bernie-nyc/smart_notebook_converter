CairoSVG with Cairo: install a wheel that bundles Cairo. On Windows run pip install "cairosvg==2.5.2"; this package contains the necessary DLLs. On macOS run brew install cairo followed by pip install cairosvg. On Debian/Ubuntu run sudo apt-get install libcairo2 then pip install cairosvg.

ImageMagick: install ImageMagick and confirm that magick is on your PATH and can convert SVG files. On Windows download the ImageMagick 7 installer (Q16 HDRI) from the official site, check “Install legacy utilities” and “Install development headers and libraries.” On macOS run brew install imagemagick; on Debian/Ubuntu run sudo apt-get install imagemagick. Test with magick input.svg output.png.

The Python cairosvg module is installed, but the shared library libcairo is missing or cannot be located, so the import succeeds but every render attempt fails and logs those warnings. To eliminate them you must install the Cairo C library and make sure it’s discoverable:

Windows: The standard pip install cairosvg relies on an external GTK/Cairo runtime. Either install GTK‑3 (which bundles the Cairo DLLs) and add its bin directory to your PATH, or install the wheel with built‑in dependencies, e.g. pip install "cairosvg==2.5.2". Confirm that libcairo-2.dll exists somewhere on your system and that its directory appears in your PATH.

macOS (Homebrew): Run brew install cairo and brew install libffi. If you’re on Apple Silicon, Homebrew installs libraries under /opt/homebrew/lib; on Intel it’s /usr/local/lib. The loader looks in /usr/local/lib by default, so create a symlink or update your DYLD_LIBRARY_PATH, for example:
ln -s /opt/homebrew/lib/libcairo.2.dylib /usr/local/lib/libcairo.2.dylib.

Linux (Debian/Ubuntu): Install Cairo and its development headers with sudo apt-get install libcairo2 libffi6. If you’re using a Python Slim image or a conda environment, set LD_LIBRARY_PATH to include /usr/lib or wherever libcairo.so.2 resides.

Fallback to ImageMagick: The script falls back to the magick command if CairoSVG fails. Make sure that ImageMagick is installed and that the magick executable is on your PATH. On Windows you should see a file named magick.exe in the install folder; add that folder to PATH. On macOS brew install imagemagick installs it in /usr/local/bin. Test with magick input.svg output.png at the command line; if it fails, your ImageMagick build lacks SVG support.

Once the Cairo library is available or the magick command is working, the warnings will disappear and the conversion will proceed correctly.
