class DocxTemplateBlockDefinition
{
    public string Name { get; }
    public string PropertyName { get; } // The field name in the parent DTO
    public bool IsRepeatable { get; }
    public List<string> Fields { get; set; } = new List<string>();
    public List<DocxTemplateBlockDefinition> Children { get; } = new List<DocxTemplateBlockDefinition>();

    public DocxTemplateBlockDefinition(string name, string propertyName, bool isRepeatable)
    {
        Name = name;
        PropertyName = propertyName;
        IsRepeatable = isRepeatable;
    }
}
