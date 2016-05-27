# vba_utilities
Reusable modules and dev utilities for the Macmillan Word template and macros.

## dependencies
Externally produced modules that must be incorporated into the main Utilities template file, but which are not part of this repo. Currently includes the wonderful [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) and [VBA-Dictonary](https://github.com/VBA-tools/VBA-Dictionary). If you clone this repo, be sure to manually add the primary modules from those projects in the `/dependencies` directory, the contents of which are not tracked here.

## Utilities
A series of modules containing classes and functions required for other Macmillan VBA projects, notably [Word-template](https://github.com/macmillanpublishers/Word-template). Just save the `MacroUtilities.dotm` file in the same directory as your primary template, be sure it is loaded as an add-in, and you can call procedures in the format `Project.Module.Procedure`.

## Dev Tools
A series of macros to help with VBA development.
