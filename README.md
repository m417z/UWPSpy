<div align="center">
	<img src="icon.png" style="width: 64px;">
	<h1>UWPSpy</h1>
</div>

An inspection tool for UWP applications. Seamlessly view and manipulate UI
elements and their properties in real time.

[🏠 Homepage](https://ramensoftware.com/uwpspy)

![Screenshot](screenshot.png)

## Usage

- Download the latest release from
  [here](https://ramensoftware.com/downloads/uwpspy.zip).
- Run `UWPSpy.exe` and select the UWP application you want to spy on.
- A window with the application's UI elements will appear for each UI thread of
  the target application.
- A UWPSpy spying window can be closed, in which case it will become hidden. It
  can be shown again by running `UWPSpy.exe` again and selecting the same
  application.
- To completely stop spying on an application, its process must be terminated.

## Demo

[![Demo video](screenshot-video.png)](https://youtu.be/Zxgk_BOVpfk)
*Click on the image to view the demo video on YouTube.*

## Supported Windows versions

UWPSpy uses the [XAML Diagnostic
APIs](https://learn.microsoft.com/en-us/windows/win32/api/xamlom/nf-xamlom-initializexamldiagnosticsex)
which were added in Windows 10, version 1703. Earlier versions of Windows are
not supported.

## References

- The
  [`ExplorerTAP`](https://github.com/TranslucentTB/TranslucentTB/tree/release/ExplorerTAP)
  part of the [TranslucentTB](https://github.com/TranslucentTB/TranslucentTB)
  project. That's the only usage example of the XAML Diagnostic APIs I could
  find on the internet.
- [Ahmed Walid (ahmed605)](https://github.com/ahmed605), thanks for helping with
  many UWP-related questions.
