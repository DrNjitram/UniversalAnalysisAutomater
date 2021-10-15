pyinstaller main.py -F -n DMA_unpacker
del /f DMA_unpacker.exe DMA_unpacker.spec
copy "%cd%\dist\DMA_unpacker.exe" "%cd%\DMA_unpacker.exe"
rd "%cd%\dist" "%cd%\build" /S /Q