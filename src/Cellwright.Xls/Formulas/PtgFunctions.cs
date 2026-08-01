using System.Collections.Frozen;

namespace Cellwright.Xls.Formulas;

/// <summary>
/// The built-in function table (Ftab, MS-XLS 2.5.198.17). A formula stores a function by index
/// rather than by name, so this table is what turns <c>0x004F</c> back into <c>ROUND</c>.
/// </summary>
/// <remarks>
/// Indexes are historical and never renumbered, which is why the table is written out in full
/// rather than derived: the gaps are as meaningful as the entries.
/// </remarks>
internal static class PtgFunctions
{
    /// <summary>Marks a function whose argument count is fixed by the table rather than the record.</summary>
    public const int VariableArgumentCount = -1;

    private static readonly (ushort Index, string Name, int Arguments)[] Table =
    [
        (0, "COUNT", VariableArgumentCount), (1, "IF", VariableArgumentCount), (2, "ISNA", 1), (3, "ISERROR", 1),
        (4, "SUM", VariableArgumentCount), (5, "AVERAGE", VariableArgumentCount), (6, "MIN", VariableArgumentCount),
        (7, "MAX", VariableArgumentCount), (8, "ROW", VariableArgumentCount), (9, "COLUMN", VariableArgumentCount),
        (10, "NA", 0), (11, "NPV", VariableArgumentCount), (12, "STDEV", VariableArgumentCount),
        (13, "DOLLAR", VariableArgumentCount), (14, "FIXED", VariableArgumentCount),
        (15, "SIN", 1), (16, "COS", 1), (17, "TAN", 1), (18, "ATAN", 1), (19, "PI", 0), (20, "SQRT", 1),
        (21, "EXP", 1), (22, "LN", 1), (23, "LOG10", 1), (24, "ABS", 1), (25, "INT", 1), (26, "SIGN", 1),
        (27, "ROUND", 2), (28, "LOOKUP", VariableArgumentCount), (29, "INDEX", VariableArgumentCount),
        (30, "REPT", 2), (31, "MID", 3), (32, "LEN", 1), (33, "VALUE", 1), (34, "TRUE", 0), (35, "FALSE", 0),
        (36, "AND", VariableArgumentCount), (37, "OR", VariableArgumentCount), (38, "NOT", 1),
        (39, "MOD", 2), (40, "DCOUNT", 3), (41, "DSUM", 3), (42, "DAVERAGE", 3), (43, "DMIN", 3), (44, "DMAX", 3),
        (45, "DSTDEV", 3), (46, "VAR", VariableArgumentCount), (47, "DVAR", 3), (48, "TEXT", 2),
        (49, "LINEST", VariableArgumentCount), (50, "TREND", VariableArgumentCount),
        (51, "LOGEST", VariableArgumentCount), (52, "GROWTH", VariableArgumentCount),
        (53, "GOTO", VariableArgumentCount), (54, "HALT", VariableArgumentCount),
        (56, "PV", VariableArgumentCount), (57, "FV", VariableArgumentCount), (58, "NPER", VariableArgumentCount),
        (59, "PMT", VariableArgumentCount), (60, "RATE", VariableArgumentCount), (61, "MIRR", 3),
        (62, "IRR", VariableArgumentCount), (63, "RAND", 0), (64, "MATCH", VariableArgumentCount),
        (65, "DATE", 3), (66, "TIME", 3), (67, "DAY", 1), (68, "MONTH", 1), (69, "YEAR", 1), (70, "WEEKDAY", VariableArgumentCount),
        (71, "HOUR", 1), (72, "MINUTE", 1), (73, "SECOND", 1), (74, "NOW", 0), (75, "AREAS", 1), (76, "ROWS", 1),
        (77, "COLUMNS", 1), (78, "OFFSET", VariableArgumentCount), (79, "ABSREF", 2), (80, "RELREF", 2),
        (81, "ARGUMENT", VariableArgumentCount), (82, "SEARCH", VariableArgumentCount), (83, "TRANSPOSE", 1),
        (84, "ERROR", VariableArgumentCount), (85, "STEP", 0), (86, "TYPE", 1), (87, "ECHO", VariableArgumentCount),
        (88, "SETNAME", VariableArgumentCount), (89, "CALLER", 0), (90, "DEREF", 1), (91, "WINDOWS", VariableArgumentCount),
        (92, "SERIES", VariableArgumentCount), (93, "DOCUMENTS", VariableArgumentCount), (94, "ACTIVECELL", 0),
        (95, "SELECTION", 0), (96, "RESULT", VariableArgumentCount), (97, "ATAN2", 2), (98, "ASIN", 1),
        (99, "ACOS", 1), (100, "CHOOSE", VariableArgumentCount), (101, "HLOOKUP", VariableArgumentCount),
        (102, "VLOOKUP", VariableArgumentCount), (103, "LINKS", VariableArgumentCount), (104, "INPUT", VariableArgumentCount),
        (105, "ISREF", 1), (106, "GETFORMULA", 1), (107, "GETNAME", VariableArgumentCount),
        (108, "SETVALUE", VariableArgumentCount), (109, "LOG", VariableArgumentCount), (110, "EXEC", VariableArgumentCount),
        (111, "CHAR", 1), (112, "LOWER", 1), (113, "UPPER", 1), (114, "PROPER", 1), (115, "LEFT", VariableArgumentCount),
        (116, "RIGHT", VariableArgumentCount), (117, "EXACT", 2), (118, "TRIM", 1), (119, "REPLACE", 4),
        (120, "SUBSTITUTE", VariableArgumentCount), (121, "CODE", 1), (122, "NAMES", VariableArgumentCount),
        (123, "DIRECTORY", VariableArgumentCount), (124, "FIND", VariableArgumentCount), (125, "CELL", VariableArgumentCount),
        (126, "ISERR", 1), (127, "ISTEXT", 1), (128, "ISNUMBER", 1), (129, "ISBLANK", 1), (130, "T", 1), (131, "N", 1),
        (132, "FOPEN", VariableArgumentCount), (133, "FCLOSE", 1), (134, "FSIZE", 1), (135, "FREADLN", 1),
        (136, "FREAD", 2), (137, "FWRITELN", 2), (138, "FWRITE", 2), (139, "FPOS", VariableArgumentCount),
        (140, "DATEVALUE", 1), (141, "TIMEVALUE", 1), (142, "SLN", 3), (143, "SYD", 4), (144, "DDB", VariableArgumentCount),
        (145, "GETDEF", VariableArgumentCount), (146, "REFTEXT", VariableArgumentCount), (147, "TEXTREF", VariableArgumentCount),
        (148, "INDIRECT", VariableArgumentCount), (149, "REGISTER", VariableArgumentCount), (150, "CALL", VariableArgumentCount),
        (151, "ADDBAR", VariableArgumentCount), (152, "ADDMENU", VariableArgumentCount), (153, "ADDCOMMAND", VariableArgumentCount),
        (154, "ENABLECOMMAND", VariableArgumentCount), (155, "CHECKCOMMAND", VariableArgumentCount),
        (156, "RENAMECOMMAND", VariableArgumentCount), (157, "SHOWBAR", VariableArgumentCount),
        (158, "DELETEMENU", VariableArgumentCount), (159, "DELETECOMMAND", VariableArgumentCount),
        (160, "GETCHARTITEM", VariableArgumentCount), (161, "DIALOGBOX", 1), (162, "CLEAN", 1),
        (163, "MDETERM", 1), (164, "MINVERSE", 1), (165, "MMULT", 2), (166, "FILES", VariableArgumentCount),
        (167, "IPMT", VariableArgumentCount), (168, "PPMT", VariableArgumentCount), (169, "COUNTA", VariableArgumentCount),
        (170, "CANCELKEY", VariableArgumentCount), (175, "INITIATE", 2), (176, "REQUEST", 2), (177, "POKE", 3),
        (178, "EXECUTE", 2), (179, "TERMINATE", 1), (180, "RESTART", VariableArgumentCount),
        (181, "HELP", VariableArgumentCount), (182, "GETBAR", VariableArgumentCount), (183, "PRODUCT", VariableArgumentCount),
        (184, "FACT", 1), (185, "GETCELL", VariableArgumentCount), (186, "GETWORKSPACE", 1),
        (187, "GETWINDOW", VariableArgumentCount), (188, "GETDOCUMENT", VariableArgumentCount),
        (189, "DPRODUCT", 3), (190, "ISNONTEXT", 1), (191, "GETNOTE", VariableArgumentCount),
        (192, "NOTE", VariableArgumentCount), (193, "STDEVP", VariableArgumentCount), (194, "VARP", VariableArgumentCount),
        (195, "DSTDEVP", 3), (196, "DVARP", 3), (197, "TRUNC", VariableArgumentCount), (198, "ISLOGICAL", 1),
        (199, "DCOUNTA", 3), (200, "DELETEBAR", 1), (201, "UNREGISTER", 1), (204, "USDOLLAR", VariableArgumentCount),
        (205, "FINDB", VariableArgumentCount), (206, "SEARCHB", VariableArgumentCount), (207, "REPLACEB", 4),
        (208, "LEFTB", VariableArgumentCount), (209, "RIGHTB", VariableArgumentCount), (210, "MIDB", 3),
        (211, "LENB", 1), (212, "ROUNDUP", 2), (213, "ROUNDDOWN", 2), (214, "ASC", 1), (215, "DBCS", 1),
        (216, "RANK", VariableArgumentCount), (219, "ADDRESS", VariableArgumentCount), (220, "DAYS360", VariableArgumentCount),
        (221, "TODAY", 0), (222, "VDB", VariableArgumentCount), (227, "MEDIAN", VariableArgumentCount),
        (228, "SUMPRODUCT", VariableArgumentCount), (229, "SINH", 1), (230, "COSH", 1), (231, "TANH", 1),
        (232, "ASINH", 1), (233, "ACOSH", 1), (234, "ATANH", 1), (235, "DGET", 3), (236, "CREATEOBJECT", VariableArgumentCount),
        (237, "VOLATILE", VariableArgumentCount), (238, "LASTERROR", 0), (239, "CUSTOMUNDO", VariableArgumentCount),
        (240, "CUSTOMREPEAT", VariableArgumentCount), (241, "FORMULACONVERT", VariableArgumentCount),
        (242, "GETLINKINFO", VariableArgumentCount), (243, "TEXTBOX", VariableArgumentCount), (244, "INFO", 1),
        (245, "GROUP", 0), (246, "GETOBJECT", VariableArgumentCount), (247, "DB", VariableArgumentCount),
        (248, "PAUSE", VariableArgumentCount), (250, "RESUME", VariableArgumentCount), (252, "FREQUENCY", 2),
        (253, "ADDTOOLBAR", VariableArgumentCount), (254, "DELETETOOLBAR", VariableArgumentCount),
        (255, "", VariableArgumentCount), // Reserved: a user-defined function reached through PtgFuncVar.
        (256, "RESETTOOLBAR", 1), (257, "EVALUATE", 1), (258, "GETTOOLBAR", VariableArgumentCount),
        (259, "GETTOOL", VariableArgumentCount), (260, "SPELLINGCHECK", VariableArgumentCount),
        (261, "ERROR.TYPE", 1), (262, "APPTITLE", VariableArgumentCount), (263, "WINDOWTITLE", VariableArgumentCount),
        (264, "SAVETOOLBAR", VariableArgumentCount), (265, "ENABLETOOL", 3), (266, "PRESSTOOL", 3),
        (267, "REGISTERID", VariableArgumentCount), (268, "GETWORKBOOK", VariableArgumentCount),
        (269, "AVEDEV", VariableArgumentCount), (270, "BETADIST", VariableArgumentCount), (271, "GAMMALN", 1),
        (272, "BETAINV", VariableArgumentCount), (273, "BINOMDIST", 4), (274, "CHIDIST", 2), (275, "CHIINV", 2),
        (276, "COMBIN", 2), (277, "CONFIDENCE", 3), (278, "CRITBINOM", 3), (279, "EVEN", 1), (280, "EXPONDIST", 3),
        (281, "FDIST", 3), (282, "FINV", 3), (283, "FISHER", 1), (284, "FISHERINV", 1), (285, "FLOOR", 2),
        (286, "GAMMADIST", 4), (287, "GAMMAINV", 3), (288, "CEILING", 2), (289, "HYPGEOMDIST", 4),
        (290, "LOGNORMDIST", 3), (291, "LOGINV", 3), (292, "NEGBINOMDIST", 3), (293, "NORMDIST", 4),
        (294, "NORMSDIST", 1), (295, "NORMINV", 3), (296, "NORMSINV", 1), (297, "STANDARDIZE", 3), (298, "ODD", 1),
        (299, "PERMUT", 2), (300, "POISSON", 3), (301, "TDIST", 3), (302, "WEIBULL", 4), (303, "SUMXMY2", 2),
        (304, "SUMX2MY2", 2), (305, "SUMX2PY2", 2), (306, "CHITEST", 2), (307, "CORREL", 2),
        (308, "COVAR", 2), (309, "FORECAST", 3), (310, "FTEST", 2), (311, "INTERCEPT", 2), (312, "PEARSON", 2),
        (313, "RSQ", 2), (314, "STEYX", 2), (315, "SLOPE", 2), (316, "TTEST", 4), (317, "PROB", VariableArgumentCount),
        (318, "DEVSQ", VariableArgumentCount), (319, "GEOMEAN", VariableArgumentCount), (320, "HARMEAN", VariableArgumentCount),
        (321, "SUMSQ", VariableArgumentCount), (322, "KURT", VariableArgumentCount), (323, "SKEW", VariableArgumentCount),
        (324, "ZTEST", VariableArgumentCount), (325, "LARGE", 2), (326, "SMALL", 2), (327, "QUARTILE", 2),
        (328, "PERCENTILE", 2), (329, "PERCENTRANK", VariableArgumentCount), (330, "MODE", VariableArgumentCount),
        (331, "TRIMMEAN", 2), (332, "TINV", 2), (334, "MOVIECOMMAND", VariableArgumentCount),
        (335, "GETMOVIE", VariableArgumentCount), (336, "CONCATENATE", VariableArgumentCount), (337, "POWER", 2),
        (338, "PIVOTADDDATA", VariableArgumentCount), (339, "GETPIVOTTABLE", VariableArgumentCount),
        (340, "GETPIVOTFIELD", VariableArgumentCount), (341, "GETPIVOTITEM", VariableArgumentCount),
        (342, "RADIANS", 1), (343, "DEGREES", 1), (344, "SUBTOTAL", VariableArgumentCount),
        (345, "SUMIF", VariableArgumentCount), (346, "COUNTIF", 2), (347, "COUNTBLANK", 1),
        (348, "SCENARIOGET", VariableArgumentCount), (349, "OPTIONSLISTSGET", 1), (350, "ISPMT", 4),
        (351, "DATEDIF", 3), (352, "DATESTRING", 1), (353, "NUMBERSTRING", 2), (354, "ROMAN", VariableArgumentCount),
        (355, "OPENDIALOG", VariableArgumentCount), (356, "SAVEDIALOG", VariableArgumentCount),
        (357, "VIEWGET", VariableArgumentCount), (358, "GETPIVOTDATA", VariableArgumentCount),
        (359, "HYPERLINK", VariableArgumentCount), (360, "PHONETIC", 1), (361, "AVERAGEA", VariableArgumentCount),
        (362, "MAXA", VariableArgumentCount), (363, "MINA", VariableArgumentCount), (364, "STDEVPA", VariableArgumentCount),
        (365, "VARPA", VariableArgumentCount), (366, "STDEVA", VariableArgumentCount), (367, "VARA", VariableArgumentCount),
        (368, "BAHTTEXT", 1), (369, "THAIDAYOFWEEK", 1), (370, "THAIDIGIT", 1), (371, "THAIMONTHOFYEAR", 1),
        (372, "THAINUMSOUND", 1), (373, "THAINUMSTRING", 1), (374, "THAISTRINGLENGTH", 1), (375, "ISTHAIDIGIT", 1),
        (376, "ROUNDBAHTDOWN", 1), (377, "ROUNDBAHTUP", 1), (378, "THAIYEAR", 1), (379, "RTD", VariableArgumentCount),
        (380, "CUBEVALUE", VariableArgumentCount), (381, "CUBEMEMBER", VariableArgumentCount),
        (382, "CUBEMEMBERPROPERTY", VariableArgumentCount), (383, "CUBERANKEDMEMBER", VariableArgumentCount),
        (384, "HEX2BIN", VariableArgumentCount), (385, "HEX2DEC", 1), (386, "HEX2OCT", VariableArgumentCount),
        (387, "DEC2BIN", VariableArgumentCount), (388, "DEC2HEX", VariableArgumentCount), (389, "DEC2OCT", VariableArgumentCount),
        (390, "OCT2BIN", VariableArgumentCount), (391, "OCT2HEX", VariableArgumentCount), (392, "OCT2DEC", 1),
        (393, "BIN2DEC", 1), (394, "BIN2OCT", VariableArgumentCount), (395, "BIN2HEX", VariableArgumentCount),
        (396, "IMSUB", 2), (397, "IMDIV", 2), (398, "IMPOWER", 2), (399, "IMABS", 1), (400, "IMSQRT", 1),
        (401, "IMLN", 1), (402, "IMLOG2", 1), (403, "IMLOG10", 1), (404, "IMSIN", 1), (405, "IMCOS", 1),
        (406, "IMEXP", 1), (407, "IMARGUMENT", 1), (408, "IMCONJUGATE", 1), (409, "IMAGINARY", 1), (410, "IMREAL", 1),
        (411, "COMPLEX", VariableArgumentCount), (412, "IMSUM", VariableArgumentCount), (413, "IMPRODUCT", VariableArgumentCount),
        (414, "SERIESSUM", 4), (415, "FACTDOUBLE", 1), (416, "SQRTPI", 1), (417, "QUOTIENT", 2), (418, "DELTA", VariableArgumentCount),
        (419, "GESTEP", VariableArgumentCount), (420, "ISEVEN", 1), (421, "ISODD", 1), (422, "MROUND", 2),
        (423, "ERF", VariableArgumentCount), (424, "ERFC", 1), (425, "BESSELJ", 2), (426, "BESSELK", 2),
        (427, "BESSELY", 2), (428, "BESSELI", 2), (429, "XIRR", VariableArgumentCount), (430, "XNPV", 3),
        (431, "PRICEMAT", VariableArgumentCount), (432, "YIELDMAT", VariableArgumentCount),
        (433, "INTRATE", VariableArgumentCount), (434, "RECEIVED", VariableArgumentCount),
        (435, "DISC", VariableArgumentCount), (436, "PRICEDISC", VariableArgumentCount),
        (437, "YIELDDISC", VariableArgumentCount), (438, "TBILLEQ", 3), (439, "TBILLPRICE", 3), (440, "TBILLYIELD", 3),
        (441, "PRICE", VariableArgumentCount), (442, "YIELD", VariableArgumentCount), (443, "DOLLARDE", 2),
        (444, "DOLLARFR", 2), (445, "NOMINAL", 2), (446, "EFFECT", 2), (447, "CUMPRINC", 6), (448, "CUMIPMT", 6),
        (449, "EDATE", 2), (450, "EOMONTH", 2), (451, "YEARFRAC", VariableArgumentCount), (452, "COUPDAYBS", VariableArgumentCount),
        (453, "COUPDAYS", VariableArgumentCount), (454, "COUPDAYSNC", VariableArgumentCount),
        (455, "COUPNCD", VariableArgumentCount), (456, "COUPNUM", VariableArgumentCount),
        (457, "COUPPCD", VariableArgumentCount), (458, "DURATION", VariableArgumentCount),
        (459, "MDURATION", VariableArgumentCount), (460, "ODDLPRICE", VariableArgumentCount),
        (461, "ODDLYIELD", VariableArgumentCount), (462, "ODDFPRICE", VariableArgumentCount),
        (463, "ODDFYIELD", VariableArgumentCount), (464, "RANDBETWEEN", 2), (465, "WEEKNUM", VariableArgumentCount),
        (466, "AMORDEGRC", VariableArgumentCount), (467, "AMORLINC", VariableArgumentCount),
        (468, "CONVERT", 3), (469, "ACCRINT", VariableArgumentCount), (470, "ACCRINTM", VariableArgumentCount),
        (471, "WORKDAY", VariableArgumentCount), (472, "NETWORKDAYS", VariableArgumentCount),
        (473, "GCD", VariableArgumentCount), (474, "MULTINOMIAL", VariableArgumentCount), (475, "LCM", VariableArgumentCount),
        (476, "FVSCHEDULE", 2), (477, "CUBEKPIMEMBER", VariableArgumentCount), (478, "CUBESET", VariableArgumentCount),
        (479, "CUBESETCOUNT", 1), (480, "IFERROR", 2), (481, "COUNTIFS", VariableArgumentCount),
        (482, "SUMIFS", VariableArgumentCount), (483, "AVERAGEIF", VariableArgumentCount),
        (484, "AVERAGEIFS", VariableArgumentCount), (485, "AGGREGATE", VariableArgumentCount),
        (486, "BINOM.DIST", 4), (487, "BINOM.INV", 3), (488, "CONFIDENCE.NORM", 3), (489, "CONFIDENCE.T", 3),
        (490, "CHISQ.TEST", 2), (491, "F.TEST", 2), (492, "COVARIANCE.P", 2), (493, "COVARIANCE.S", 2),
        (494, "EXPON.DIST", 3), (495, "GAMMA.DIST", 4), (496, "GAMMA.INV", 3), (497, "MODE.MULT", VariableArgumentCount),
        (498, "MODE.SNGL", VariableArgumentCount), (499, "NORM.DIST", 4), (500, "NORM.INV", 3),
        (501, "PERCENTILE.EXC", 2), (502, "PERCENTILE.INC", 2), (503, "PERCENTRANK.EXC", VariableArgumentCount),
        (504, "PERCENTRANK.INC", VariableArgumentCount), (505, "POISSON.DIST", 3), (506, "QUARTILE.EXC", 2),
        (507, "QUARTILE.INC", 2), (508, "RANK.AVG", VariableArgumentCount), (509, "RANK.EQ", VariableArgumentCount),
        (510, "STDEV.S", VariableArgumentCount), (511, "STDEV.P", VariableArgumentCount), (512, "T.DIST", 3),
        (513, "T.DIST.2T", 2), (514, "T.DIST.RT", 2), (515, "T.INV", 2), (516, "T.INV.2T", 2), (517, "VAR.S", VariableArgumentCount),
        (518, "VAR.P", VariableArgumentCount), (519, "WEIBULL.DIST", 4), (520, "NETWORKDAYS.INTL", VariableArgumentCount),
        (521, "WORKDAY.INTL", VariableArgumentCount), (522, "ECMA.CEILING", 2), (523, "ISO.CEILING", VariableArgumentCount),
        (525, "BETA.DIST", VariableArgumentCount), (526, "BETA.INV", VariableArgumentCount),
        (527, "CHISQ.DIST", 3), (528, "CHISQ.DIST.RT", 2), (529, "CHISQ.INV", 2), (530, "CHISQ.INV.RT", 2),
        (531, "F.DIST", 4), (532, "F.DIST.RT", 3), (533, "F.INV", 3), (534, "F.INV.RT", 3), (535, "HYPGEOM.DIST", 5),
        (536, "LOGNORM.DIST", 4), (537, "LOGNORM.INV", 3), (538, "NEGBINOM.DIST", 4), (539, "NORM.S.DIST", 2),
        (540, "NORM.S.INV", 1), (541, "T.TEST", 4), (542, "Z.TEST", VariableArgumentCount), (543, "ERF.PRECISE", 1),
        (544, "ERFC.PRECISE", 1), (545, "GAMMALN.PRECISE", 1), (546, "CEILING.PRECISE", VariableArgumentCount),
        (547, "FLOOR.PRECISE", VariableArgumentCount),
    ];

    private static readonly FrozenDictionary<ushort, string> NameByIndex =
        Table.ToFrozenDictionary(static entry => entry.Index, static entry => entry.Name);

    private static readonly FrozenDictionary<ushort, int> ArgumentsByIndex =
        Table.ToFrozenDictionary(static entry => entry.Index, static entry => entry.Arguments);

    private static readonly FrozenDictionary<string, ushort> IndexByName = BuildIndexByName();

    private static FrozenDictionary<string, ushort> BuildIndexByName()
    {
        var map = new Dictionary<string, ushort>(Table.Length, StringComparer.OrdinalIgnoreCase);
        foreach ((ushort index, string name, _) in Table)
        {
            if (name.Length > 0)
                map.TryAdd(name, index);
        }

        return map.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Returns the name of a built-in function, or <see langword="null"/> for an unknown index.</summary>
    public static string? GetName(ushort index) =>
        NameByIndex.TryGetValue(index, out string? name) && name.Length > 0 ? name : null;

    /// <summary>Returns how many arguments a fixed-arity function takes, or -1 when it varies.</summary>
    public static int GetArgumentCount(ushort index) =>
        ArgumentsByIndex.TryGetValue(index, out int count) ? count : VariableArgumentCount;

    /// <summary>Returns the index of a function by name, or <see langword="null"/> when it is not built in.</summary>
    public static ushort? GetIndex(string name) =>
        IndexByName.TryGetValue(name, out ushort index) ? index : null;
}
