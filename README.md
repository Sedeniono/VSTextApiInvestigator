# VSTextApiInvestigator <!-- omit in toc -->

## Introduction
A Visual Studio 2022 extension that serves as a playground for calling Visual Studio APIs that can be used to get information about source code structure/syntax/elements.
As such, it is not an extension that is useful to end-users.
Instead, I used it to figure out the behavior of Visual Studio APIs while writing other extensions (such as [VSDoxyHighlighter](https://github.com/Sedeniono/VSDoxyHighlighter)).

## Using the extension

Clone the repository, open the Visual Studio solution `VSTextApiInvestigator.sln` and build. You can then start debugging (causing VS to deploy the built vsix-package to the VS experimental instance).
There is no downloadable vsix package provided because the intention is to play around with the extension's source code.

The extension adds a new tool window in the menu "View" &rarr; "Other Windows" &rarr; "VSTextApiInvestigator tool window".
Once opened, the window is updated whenever the user changes the caret or selection in a text view.

The window currently can show information from the following APIs:
* [`ITextStructureNavigator`](https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.text.operations.itextstructurenavigator?view=visualstudiosdk-2022)
* [`EditPoint.CodeElement`](https://learn.microsoft.com/en-us/dotnet/api/envdte.editpoint.codeelement?view=visualstudiosdk-2022) (also compare [CodeModel](https://learn.microsoft.com/en-us/dotnet/api/envdte.codemodel?view=visualstudiosdk-2022) and [FileCodeModel](https://learn.microsoft.com/en-us/dotnet/api/envdte.filecodemodel?view=visualstudiosdk-2022)).



