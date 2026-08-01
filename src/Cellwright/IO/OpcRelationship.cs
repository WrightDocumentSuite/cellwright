namespace Cellwright.IO;

/// <summary>
/// A single OPC relationship: how one package part points at another (or at an external resource).
/// </summary>
/// <param name="Id">Relationship identifier unique within the source part, e.g. <c>rId1</c>.</param>
/// <param name="Type">Relationship type URI.</param>
/// <param name="Target">Target part path or external URI.</param>
/// <param name="IsExternal">Whether the target lives outside the package.</param>
internal readonly record struct OpcRelationship(string Id, string Type, string Target, bool IsExternal = false);
