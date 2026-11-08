; Unshipped analyzer release
; https://github.com/dotnet/roslyn-analyzers/blob/main/src/Microsoft.CodeAnalysis.Analyzers/ReleaseTrackingAnalyzers.Help.md

### New Rules

Rule ID | Category           | Severity | Notes
--------|--------------------|----------|------------------------------------------------------------
CWM001  | Cellwright.Mapping | Error    | A type marked with [ExcelRow] and all its containing types must be declared partial.
CWM002  | Cellwright.Mapping | Error    | A type shape cannot implement a generated row schema.
CWM003  | Cellwright.Mapping | Error    | Generic types cannot receive a generated row schema.
CWM004  | Cellwright.Mapping | Error    | A type marked with [ExcelRow] must have at least one mappable column.
CWM005  | Cellwright.Mapping | Error    | A primary constructor parameter must have a matching property.
CWM006  | Cellwright.Mapping | Error    | A primary constructor parameter cannot be ignored with [ExcelIgnore].
CWM007  | Cellwright.Mapping | Error    | A mapped property has an unsupported type.
