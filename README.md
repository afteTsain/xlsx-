# xlsx


##为Power Query数据汇总做好清洗工作

= Table.Combine(Table.TransformColumns(Folder.Files("D:\测试数据"),{"Content",each Table.PromoteHeaders(Table.Skip(Excel.Workbook(_)[Data]{0},3))})[Content])

= Table.Combine(Table.TransformColumns(源,{"Content",each Table.PromoteHeaders(Table.Skip(Excel.Workbook(_)[Data]{0},3))})[Content])

=Table.Combine(List.Combine(List.Transform(源[Content],each List.Transform(Excel.Workbook(_)[Data],(x)=>Table.PromoteHeaders(Table.Skip(Excel.Workbook(_)[Data]{0},3))))))
