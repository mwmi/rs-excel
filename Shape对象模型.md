Excel 里的 `Shape` 对象模型是操作图形、图片、文本框、按钮、连接线、SmartArt、图表等的核心。下面按“实际使用频率 + 实战推荐”给你系统整理。

---

# 一、Shape 对象模型层级

```text
Worksheet / Chart
  └─ Shapes (ShapeRange / Shape 集合)
       ├─ Shape (单个形状)
       │    ├─ Fill / Line / Shadow / Glow / SoftEdge / Reflection
       │    ├─ TextFrame / TextFrame2
       │    ├─ Adjustments
       │    ├─ GroupItems (组合形状的子形状)
       │    └─ ...
       └─ ShapeRange (多个形状的集合)
```

核心对象：

| 对象             | 说明                   |
| ---------------- | ---------------------- |
| `Shapes`       | 工作表上所有形状的集合 |
| `Shape`        | 单个形状               |
| `ShapeRange`   | 多个形状组成的范围     |
| `GroupShapes`  | 组合形状的子形状集合   |
| `TextFrame2`   | 文本框架（新版）       |
| `FillFormat`   | 填充格式               |
| `LineFormat`   | 线条格式               |
| `ShadowFormat` | 阴影格式               |
| `Adjustments`  | 调整值集合             |

---

# 二、Shapes 集合常用 API

## 1. 常用属性

| 属性       | 作用             | 示例                                           |
| ---------- | ---------------- | ---------------------------------------------- |
| `Count`  | 形状数量         | `ws.Shapes.Count`                            |
| `Item`   | 按索引或名称引用 | `ws.Shapes.Item(1)`、`ws.Shapes("矩形 1")` |
| `Range`  | 返回 ShapeRange  | `ws.Shapes.Range(Array("矩形 1", "矩形 2"))` |
| `Parent` | 父对象           | `ws.Shapes.Parent`                           |

## 2. 常用方法

| 方法              | 作用          | 示例                                                                      |
| ----------------- | ------------- | ------------------------------------------------------------------------- |
| `AddShape`      | 添加自选图形  | `ws.Shapes.AddShape(msoShapeRectangle, 100, 100, 200, 50)`              |
| `AddTextbox`    | 添加文本框    | `ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 50)` |
| `AddLine`       | 添加直线      | `ws.Shapes.AddLine(0, 0, 100, 100)`                                     |
| `AddConnector`  | 添加连接线    | `ws.Shapes.AddConnector(msoConnectorStraight, 0, 0, 100, 100)`          |
| `AddPicture`    | 添加图片      | `ws.Shapes.AddPicture("C:\a.jpg", msoFalse, msoTrue, 100, 100)`         |
| `AddOLEObject`  | 添加 OLE 对象 | `ws.Shapes.AddOLEObject(...)`                                           |
| `AddChart2`     | 添加图表      | `ws.Shapes.AddChart2(240, xlColumnClustered)`                           |
| `AddSmartArt`   | 添加 SmartArt | `ws.Shapes.AddSmartArt(...)`                                            |
| `AddCurve`      | 添加曲线      | `ws.Shapes.AddCurve(...)`                                               |
| `AddPolyline`   | 添加多段线    | `ws.Shapes.AddPolyline(...)`                                            |
| `AddCallout`    | 添加标注      | `ws.Shapes.AddCallout(...)`                                             |
| `AddLabel`      | 添加标签      | `ws.Shapes.AddLabel(...)`                                               |
| `SelectAll`     | 全选形状      | `ws.Shapes.SelectAll`                                                   |
| `BuildFreeform` | 构建自由形状  | `ws.Shapes.BuildFreeform(...)`                                          |

## 3. 遍历与批量操作

```vba
' 遍历所有形状
Dim shp As Shape
For Each shp In ws.Shapes
    Debug.Print shp.Name, shp.Type
Next shp

' 遍历指定类型
For Each shp In ws.Shapes
    If shp.Type = msoShapeRectangle Then
        ' 处理矩形
    End If
Next shp

' 批量选择
ws.Shapes.Range(Array("矩形 1", "矩形 2")).Select

' 批量设置属性
Dim sr As ShapeRange
Set sr = ws.Shapes.Range(Array("矩形 1", "矩形 2"))
sr.Fill.ForeColor.RGB = RGB(255, 0, 0)
```

---

# 三、Shape 对象常用 API

## 1. 标识与类型

| 属性                | 作用         | 示例                                      |
| ------------------- | ------------ | ----------------------------------------- |
| `Name`            | 名称         | `shp.Name = "myRect"`                   |
| `Type`            | 类型枚举     | `shp.Type = msoPicture`                 |
| `AutoShapeType`   | 自选图形类型 | `shp.AutoShapeType = msoShapeRectangle` |
| `ID`              | 唯一 ID      | `shp.ID`                                |
| `Parent`          | 父对象       | `shp.Parent`                            |
| `ParentGroup`     | 所属组合     | `shp.ParentGroup`                       |
| `TopLeftCell`     | 左上角单元格 | `shp.TopLeftCell.Address`               |
| `BottomRightCell` | 右下角单元格 | `shp.BottomRightCell.Address`           |

常用 `Type` 枚举：

| 枚举                      | 说明       |
| ------------------------- | ---------- |
| `msoShapeRectangle`     | 矩形       |
| `msoShapeOval`          | 椭圆       |
| `msoShapeRightTriangle` | 直角三角形 |
| `msoTextBox`            | 文本框     |
| `msoPicture`            | 图片       |
| `msoChart`              | 图表       |
| `msoLine`               | 直线       |
| `msoConnector`          | 连接线     |
| `msoGroup`              | 组合       |
| `msoSmartArt`           | SmartArt   |
| `msoAutoShape`          | 自选图形   |
| `msoFreeform`           | 自由形状   |
| `msoCallout`            | 标注       |
| `msoComment`            | 批注       |

## 2. 位置与大小

| 属性                           | 作用           | 示例                                                    |
| ------------------------------ | -------------- | ------------------------------------------------------- |
| `Left`                       | 左边距         | `shp.Left = 100`                                      |
| `Top`                        | 上边距         | `shp.Top = 100`                                       |
| `Width`                      | 宽度           | `shp.Width = 200`                                     |
| `Height`                     | 高度           | `shp.Height = 50`                                     |
| `Rotation`                   | 旋转角度       | `shp.Rotation = 45`                                   |
| `Flip`                       | 翻转           | `shp.Flip msoFlipHorizontal`                          |
| `LockAspectRatio`            | 锁定纵横比     | `shp.LockAspectRatio = msoTrue`                       |
| `Placement`                  | 随单元格变化   | `shp.Placement = xlMoveAndSize`                       |
| `RelativeHorizontalPosition` | 水平相对位置   | `shp.RelativeHorizontalPosition = xlRelativePosition` |
| `RelativeVerticalPosition`   | 垂直相对位置   | `shp.RelativeVerticalPosition = xlRelativePosition`   |
| `ScaleHeight`                | 按比例缩放高度 | `shp.ScaleHeight 1.5, msoTrue`                        |
| `ScaleWidth`                 | 按比例缩放宽度 | `shp.ScaleWidth 1.5, msoTrue`                         |
| `Locked`                     | 保护时锁定     | `shp.Locked = msoTrue`                                |

常用 `Placement` 枚举：

| 枚举                  | 说明                     |
| --------------------- | ------------------------ |
| `xlMoveAndSize`     | 随单元格移动并调整大小   |
| `xlMoveButDontSize` | 随单元格移动，不调整大小 |
| `xlFreeFloating`    | 自由浮动                 |
| `xlMoveAndSize`     | 随单元格移动和调整大小   |

## 3. 显示与外观

| 属性                | 作用       | 示例                                  |
| ------------------- | ---------- | ------------------------------------- |
| `Visible`         | 是否可见   | `shp.Visible = msoFalse`            |
| `Fill`            | 填充       | `shp.Fill.ForeColor.RGB = vbRed`    |
| `Line`            | 线条       | `shp.Line.Visible = msoTrue`        |
| `Shadow`          | 阴影       | `shp.Shadow.Visible = msoTrue`      |
| `Glow`            | 发光       | `shp.Glow.Radius = 10`              |
| `SoftEdge`        | 柔化边缘   | `shp.SoftEdge.Radius = 5`           |
| `Reflection`      | 映像       | `shp.Reflection.Visible = msoTrue`  |
| `ThreeD`          | 三维格式   | `shp.ThreeD.RotationX = 30`         |
| `TextEffect`      | 艺术字效果 | `shp.TextEffect.FontName = "Arial"` |
| `AlternativeText` | 替代文本   | `shp.AlternativeText = "说明"`      |
| `Title`           | 标题       | `shp.Title = "标题"`                |

## 4. 层级与排列

| 属性/方法             | 作用       | 示例                                                  |
| --------------------- | ---------- | ----------------------------------------------------- |
| `ZOrderPosition`    | Z 顺序位置 | `shp.ZOrderPosition`                                |
| `ZOrder`            | 调整层级   | `shp.ZOrder msoBringToFront`                        |
| `IncrementLeft`     | 水平移动   | `shp.IncrementLeft 10`                              |
| `IncrementTop`      | 垂直移动   | `shp.IncrementTop -10`                              |
| `IncrementRotation` | 增量旋转   | `shp.IncrementRotation 15`                          |
| `Align`             | 对齐       | `shp.Align msoAlignCenters, msoTrue`                |
| `Distribute`        | 分布       | `sr.Distribute msoDistributeHorizontally, msoFalse` |

常用 ZOrder 枚举：

| 枚举                | 说明     |
| ------------------- | -------- |
| `msoBringToFront` | 置于顶层 |
| `msoSendToBack`   | 置于底层 |
| `msoBringForward` | 上移一层 |
| `msoSendBackward` | 下移一层 |

## 5. 操作与生命周期

| 方法                             | 作用         | 示例                               |
| -------------------------------- | ------------ | ---------------------------------- |
| `Delete`                       | 删除         | `shp.Delete`                     |
| `Copy`                         | 复制到剪贴板 | `shp.Copy`                       |
| `Cut`                          | 剪切         | `shp.Cut`                        |
| `Duplicate`                    | 创建副本     | `Set newShp = shp.Duplicate`     |
| `Select`                       | 选中         | `shp.Select`                     |
| `CopyPicture`                  | 复制为图片   | `shp.CopyPicture`                |
| `Group`                        | 组合         | `sr.Group`                       |
| `Ungroup`                      | 取消组合     | `shp.Ungroup`                    |
| `Regroup`                      | 重新组合     | `sr.Regroup`                     |
| `Apply`                        | 应用格式     | `shp.Apply`                      |
| `PickUp`                       | 拾取格式     | `shp.PickUp`                     |
| `SetShapesDefaultProperties`   | 设为默认属性 | `shp.SetShapesDefaultProperties` |
| `RerouteConnections`           | 重排连接线   | `shp.RerouteConnections`         |
| `ScaleHeight` / `ScaleWidth` | 缩放         | `shp.ScaleHeight 1.2, msoTrue`   |
| `Flip`                         | 翻转         | `shp.Flip msoFlipVertical`       |

## 6. 交互

| 属性          | 作用         | 示例                                        |
| ------------- | ------------ | ------------------------------------------- |
| `OnAction`  | 点击运行宏   | `shp.OnAction = "MyMacro"`                |
| `Hyperlink` | 超链接       | `shp.Hyperlink.Address = "https://..."`   |
| `HasChart`  | 是否包含图表 | `If shp.HasChart Then`                    |
| `Chart`     | 图表对象     | `shp.Chart.ChartType = xlColumnClustered` |

> `OnAction` 只能指定标准模块中的宏名，不能直接传参。

---

# 四、常用枚举速查

| 类别           | 枚举                                              | 说明        |
| -------------- | ------------------------------------------------- | ----------- |
| Shape Type     | `msoAutoShape`                                  | 自选图形    |
|                | `msoTextBox`                                    | 文本框      |
|                | `msoPicture`                                    | 图片        |
|                | `msoChart`                                      | 图表        |
|                | `msoLine`                                       | 直线        |
|                | `msoGroup`                                      | 组合        |
|                | `msoSmartArt`                                   | SmartArt    |
|                | `msoFreeform`                                   | 自由形状    |
|                | `msoConnector`                                  | 连接线      |
| AutoShapeType  | `msoShapeRectangle`                             | 矩形        |
|                | `msoShapeOval`                                  | 椭圆        |
|                | `msoShapeRoundedRectangle`                      | 圆角矩形    |
|                | `msoShapeIsoscelesTriangle`                     | 等腰三角形  |
|                | `msoShapeRightArrow`                            | 右箭头      |
|                | `msoShapeChevron`                               | 箭头/人字形 |
| Fill Type      | `msoFillSolid`                                  | 纯色        |
|                | `msoFillPatterned`                              | 图案        |
|                | `msoFillGradient`                               | 渐变        |
|                | `msoFillTextured`                               | 纹理        |
|                | `msoFillBackground`                             | 背景        |
| Line DashStyle | `msoLineSolid`                                  | 实线        |
|                | `msoLineDash`                                   | 虚线        |
|                | `msoLineDot`                                    | 点线        |
|                | `msoLineDashDot`                                | 点划线      |
| Alignment      | `msoAlignLefts`                                 | 左对齐      |
|                | `msoAlignCenters`                               | 居中对齐    |
|                | `msoAlignRights`                                | 右对齐      |
|                | `msoAlignTops`                                  | 顶对齐      |
|                | `msoAlignMiddles`                               | 居中        |
|                | `msoAlignBottoms`                               | 底对齐      |
| Distribute     | `msoDistributeHorizontally`                     | 水平分布    |
|                | `msoDistributeVertically`                       | 垂直分布    |
| TriState       | `msoTrue` / `msoFalse` / `msoTriStateMixed` | 三态        |

---

# 五、常用实战示例

## 1. 创建矩形并设置样式

```vba
Sub AddRectangle()
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 100, 100, 200, 80)

    With shp
        .Name = "myRect"
        .Fill.ForeColor.RGB = RGB(68, 114, 196)
        .Fill.Transparency = 0.1
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1.5
        .Shadow.Visible = msoTrue

        .TextFrame2.TextRange.Text = "点击我"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle

        .OnAction = "ShapeClicked"
        .AlternativeText = "自定义矩形"
    End With
End Sub

Sub ShapeClicked()
    MsgBox "你点击了矩形"
End Sub
```

## 2. 批量修改所有矩形颜色

```vba
Sub RecolorAllRectangles()
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.AutoShapeType = msoShapeRectangle Then
            shp.Fill.ForeColor.RGB = RGB(255, 192, 0)
        End If
    Next shp
End Sub
```

## 3. 对齐与分布

```vba
Sub AlignShapes()
    Dim sr As ShapeRange
    Set sr = ws.Shapes.Range(Array("矩形 1", "矩形 2", "矩形 3"))

    sr.Align msoAlignCenters, msoFalse
    sr.Distribute msoDistributeHorizontally, msoFalse
End Sub
```

## 4. 组合与取消组合

```vba
Sub GroupShapes()
    Dim sr As ShapeRange
    Dim grp As Shape

    Set sr = ws.Shapes.Range(Array("矩形 1", "矩形 2"))
    Set grp = sr.Group
    grp.Name = "myGroup"

    ' 取消组合
    ' grp.Ungroup
End Sub
```

## 5. 添加连接线

```vba
Sub AddConnector()
    Dim conn As Shape
    Set conn = ws.Shapes.AddConnector( _
        Type:=msoConnectorStraight, _
        BeginX:=100, BeginY:=100, _
        EndX:=300, EndY:=200)

    conn.Line.ForeColor.RGB = RGB(0, 0, 0)
    conn.Line.Weight = 2
End Sub
```

## 6. 删除所有形状

```vba
Sub DeleteAllShapes()
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub
```

> 删除时倒序循环，避免索引错乱。

## 7. 导出形状为图片

```vba
Sub ExportShape()
    Dim ch As Chart
    Set ch = ws.Shapes.AddChart2(240, xlColumnClustered).Chart

    ws.Shapes("myRect").CopyPicture
    ch.Paste
    ch.Export "C:\Temp\shape.png"
    ch.Parent.Delete
End Sub
```

## 8. 查找并高亮形状

```vba
Sub FindShapeByName()
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("myRect")
    On Error GoTo 0

    If Not shp Is Nothing Then
        shp.Fill.ForeColor.RGB = vbYellow
        shp.Line.ForeColor.RGB = vbRed
        shp.Line.Weight = 3
    End If
End Sub
```

---

# 六、FillFormat 常用成员

```vba
With shp.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 0, 0)
    .BackColor.RGB = RGB(0, 255, 0)
    .Transparency = 0.2
    .Solid                                    ' 纯色
    .Patterned msoPatternDarkHorizontal        ' 图案
    .GradientAngle = 45                        ' 渐变角度
    .TwoColorGradient msoGradientHorizontal, 1 ' 双色渐变
    .OneColorGradient msoGradientVertical, 1, 0.5
    .PresetTextured msoTextureCanvas           ' 纹理
    .UserPicture "C:\bg.jpg"                   ' 图片填充
    .UserTextured "C:\bg.jpg"
End With
```

---

# 七、LineFormat 常用成员

```vba
With shp.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(0, 0, 0)
    .BackColor.RGB = vbWhite
    .Weight = 1.5
    .DashStyle = msoLineDash
    .Style = msoLineSingle
    .Transparency = 0
    .BeginArrowheadStyle = msoArrowheadTriangle
    .EndArrowheadStyle = msoArrowheadStealth
    .BeginArrowheadLength = msoArrowheadShort
    .EndArrowheadWidth = msoArrowheadWide
    .JoinStyle = msoLineJoinRound
End With
```

---

# 八、TextFrame2 常用成员（针对含文本的形状）

```vba
With shp.TextFrame2
    .TextRange.Text = "Hello"
    .TextRange.Font.Name = "微软雅黑"
    .TextRange.Font.Size = 12
    .TextRange.Font.Bold = msoTrue
    .TextRange.Font.Fill.ForeColor.RGB = vbWhite
    .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    .VerticalAnchor = msoAnchorMiddle
    .HorizontalAnchor = msoAnchorCenter
    .WordWrap = msoTrue
    .AutoSize = msoAutoSizeShapeToFitText
    .MarginLeft = 5
    .MarginTop = 5
    .Orientation = msoTextOrientationHorizontal
End With
```

---

# 九、常见坑与最佳实践

1. **Shapes 是工作表级集合**：每个工作表有独立的 `Shapes`，不能跨表统一处理。
2. **名称引用优先于索引**：`Shapes(1)` 会因增删而变，尽量用 `Shapes("名称")`。
3. **删除时倒序循环**：避免索引错乱。
4. **组合形状的子形状要用 `GroupItems`**：`shp.GroupItems.Count`。
5. **组合后原形状引用可能失效**：需要重新引用。
6. **修改多个属性时用 `With`**：减少引用开销。
7. **批量操作关闭屏幕刷新**：
   ```vba
   Application.ScreenUpdating = False
   ' ...
   Application.ScreenUpdating = True
   ```
8. **OnAction 宏必须在标准模块**：不能写在 Sheet 或 ThisWorkbook 模块。
9. **图片插入位置**：`AddPicture` 的 `Left`、`Top` 以磅为单位（1 磅 ≈ 1/72 英寸）。
10. **图表是 Shape 的子对象**：`shp.HasChart` 判断，`shp.Chart` 访问。
11. **SmartArt 处理复杂**：通过 `shp.SmartArt` 访问，节点操作用 `SmartArt.Nodes`。
12. **保护工作表后**：`Locked = msoTrue` 的形状不可编辑。
13. **形状透明度和颜色**：`RGB` 函数返回值可直接赋给 `ForeColor.RGB`。
14. **不要用 Win32 API 操作 Shape**：Shape 没有独立窗口句柄，Win32 API 基本无用。
15. **形状太多会影响性能**：大量形状建议用图片或图表替代。

---

# 十、最推荐记住的 TOP 25

## Shapes 集合

1. `Count`
2. `Item`
3. `Range`
4. `AddShape`
5. `AddTextbox`
6. `AddPicture`
7. `AddChart2`
8. `AddConnector`
9. `SelectAll`

## Shape 对象

10. `Name`
11. `Type`
12. `AutoShapeType`
13. `Left` / `Top` / `Width` / `Height`
14. `Rotation`
15. `Visible`
16. `Fill`
17. `Line`
18. `Shadow`
19. `TextFrame2`
20. `OnAction`
21. `AlternativeText`
22. `Placement`
23. `ZOrder`
24. `Delete` / `Duplicate` / `Copy`
25. `Group` / `Ungroup`

---

# 十一、按场景推荐的 API

| 场景     | 推荐 API                                                                                          |
| -------- | ------------------------------------------------------------------------------------------------- |
| 创建形状 | `Shapes.AddShape`、`AddTextbox`、`AddPicture`、`AddLine`、`AddConnector`、`AddChart2` |
| 引用形状 | `Shapes("名称")`、`Shapes.Range(Array(...))`、`Shapes.Item`                                 |
| 遍历     | `For Each shp In ws.Shapes`                                                                     |
| 判断类型 | `shp.Type`、`shp.AutoShapeType`、`shp.HasChart`                                             |
| 位置大小 | `Left`、`Top`、`Width`、`Height`、`Rotation`、`Placement`                             |
| 外观     | `Fill`、`Line`、`Shadow`、`Glow`、`SoftEdge`、`Reflection`                            |
| 文本     | `TextFrame2.TextRange`、`TextFrame.Characters`                                                |
| 层级     | `ZOrder`、`ZOrderPosition`                                                                    |
| 对齐分布 | `Align`、`Distribute`、`IncrementLeft/Top`                                                  |
| 交互     | `OnAction`、`Hyperlink`                                                                       |
| 组合     | `Group`、`Ungroup`、`Regroup`、`GroupItems`                                               |
| 删除复制 | `Delete`、`Duplicate`、`Copy`、`Cut`                                                      |
| 图表     | `HasChart`、`Chart`                                                                           |
| 批量操作 | `ShapeRange`、`Shapes.Range`、`SelectAll`                                                   |

---

总结一句：
**Excel Shape 对象模型的核心就是 `Shapes` 集合 + `Shape` 对象 + `ShapeRange` 范围。**
日常最高频的是：创建（`AddShape` / `AddTextbox` / `AddPicture`）、引用（名称 / Range）、位置大小（`Left/Top/Width/Height`）、外观（`Fill/Line/Shadow`）、文本（`TextFrame2`）、层级（`ZOrder`）、交互（`OnAction`）、组合（`Group/Ungroup`）、遍历与批量删除。
优先用名称引用，遍历时判断 `Type`，批量操作关闭屏幕刷新，删除时倒序循环。
