Public Class GameForm
    Private Declare Function ReleaseCapture Lib "user32" () As Integer
    Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As VariantType) As Integer

    Private Const TitleHeight As Integer = 70 '标题栏高度
    Private Const MarginSize As Integer = 2 '卡片之间的间距
    Private Const PaddingSize As Integer = 30 '裁片集中区域与窗体边框的距离
    Private Const CardSize As Integer = 25 '卡片尺寸
    Dim ScoreList() As Integer = {0, 10, 30, 60, 100, 150, 210}
    Dim EmptyPoint As Point = New Point(-1, -1)
    Dim MousePointInLabel As Point '用于记录鼠标拖动标签时鼠标坐标与标签起点差值
    Dim Score As Integer = 0 '分数
    Dim Moved As Boolean '定义一个标识，记录是否发生了移动，以确定操作是否有效
    Dim BlankColor As Color = Color.FromArgb(100, Color.Gray)
    Dim ObjectLabelLocation() As Point
    Dim CardData(9, 9) As Boolean
    Dim ColorData(9, 9) As Color
    Dim ObjectType(2) As Integer '记录新产生的三个物体类型在 ObjectModel 里的标识
    Dim ObjectColor(2) As Color '记录新产生的三个物体的颜色
    Dim ObjectCount As Integer '还没有被放入游戏区的物体个数
    Dim ObjectLabels() As Label
    Dim ObjectModel(,) As Point = {
        {New Point(0, 0), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'单个方格
        {New Point(0, 0), New Point(1, 0), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'垂直排列的两个方格
        {New Point(0, 0), New Point(0, 1), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'水平排量的两个方格
        {New Point(1, 0), New Point(0, 1), New Point(1, 1), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第二象限的三个方格
        {New Point(0, 0), New Point(1, 0), New Point(1, 1), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第一象限的三个方格
        {New Point(0, 0), New Point(0, 1), New Point(1, 1), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第三象限的三个方格
        {New Point(0, 0), New Point(1, 0), New Point(0, 1), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第四象限的三个方格
        {New Point(0, 0), New Point(1, 0), New Point(2, 0), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'垂直的三个方格
        {New Point(0, 0), New Point(0, 1), New Point(0, 2), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'水平的三个方格
        {New Point(0, 0), New Point(1, 0), New Point(0, 1), New Point(1, 1), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'四个紧凑的方格
        {New Point(0, 0), New Point(1, 0), New Point(2, 0), New Point(3, 0), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'垂直的四个方格
        {New Point(0, 0), New Point(0, 1), New Point(0, 2), New Point(0, 3), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'水平的四个方格
        {New Point(0, 2), New Point(1, 2), New Point(2, 2), New Point(2, 1), New Point(2, 0), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第二象限的五个方格
        {New Point(0, 0), New Point(1, 0), New Point(2, 0), New Point(2, 1), New Point(2, 2), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第一象限的五个方格
        {New Point(0, 0), New Point(0, 1), New Point(0, 2), New Point(1, 2), New Point(2, 2), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第三象限的五个方格
        {New Point(0, 0), New Point(0, 1), New Point(0, 2), New Point(1, 0), New Point(2, 0), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'缺少第四象限的五个方格
        {New Point(0, 0), New Point(1, 0), New Point(2, 0), New Point(3, 0), New Point(4, 0), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'垂直的五个方格
        {New Point(0, 0), New Point(0, 1), New Point(0, 2), New Point(0, 3), New Point(0, 4), EmptyPoint, EmptyPoint, EmptyPoint, EmptyPoint},'水平的五个方格
        {New Point(0, 0), New Point(0, 1), New Point(0, 2), New Point(1, 0), New Point(1, 1), New Point(1, 2), New Point(2, 0), New Point(2, 1), New Point(2, 2)}'九个方格组成的超大正方形
    }

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ObjectLabelLocation = {ObjectLabel0.Location, ObjectLabel1.Location, ObjectLabel2.Location}
        ObjectLabels = {ObjectLabel0, ObjectLabel1, ObjectLabel2}

        For Index As Integer = 0 To 100
            CreateNewObject()
        Next

        DrawForm()
    End Sub

    ''' <summary>
    ''' 根据 CardData 数组刷新界面
    ''' </summary>
    Private Sub DrawForm()
        ScoreLabel.Text = Score.ToString
        Dim UnityBitmap As Bitmap = My.Resources._1010Resource.Background
        Using UnityGraphics As Graphics = Graphics.FromImage(UnityBitmap)
            For IndexY As Integer = 0 To 9
                For IndexX As Integer = 0 To 9
                    UnityGraphics.FillRectangle(New SolidBrush(IIf(CardData(IndexY, IndexX), ColorData(IndexY, IndexX), BlankColor)), New RectangleF(PaddingSize + IndexX * (CardSize + MarginSize), PaddingSize + IndexY * (CardSize + MarginSize) + TitleHeight, CardSize, CardSize))
                    'UnityGraphics.DrawString(IndexX & "," & IndexY, Me.Font, Brushes.Red, PaddingSize + IndexX * (CardSize + MarginSize), PaddingSize + IndexY * (CardSize + MarginSize) + TitleHeight)
                Next
            Next
        End Using
        Me.BackgroundImage = UnityBitmap
        GC.Collect() '回收内存
    End Sub

    ''' <summary>
    ''' 允许鼠标拖动窗体
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub GameForm_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown, ScoreLabel.MouseDown
        ReleaseCapture()
        SendMessageA(Me.Handle, &HA1, 2, 0&)
    End Sub

    ''' <summary>
    ''' 产生3个新的物体用于放置到游戏区中
    ''' </summary>
    Private Sub CreateNewObject()
        ObjectCount = 3
        ObjectLabel0.Show()
        ObjectLabel1.Show()
        ObjectLabel2.Show()
        For Index As Integer = 0 To 2
            ObjectType(Index) = VBMath.Rnd * 18
            Debug.Print(ObjectType(Index))
            ObjectColor(Index) = Color.FromArgb(255, VBMath.Rnd * 255, VBMath.Rnd * 255, VBMath.Rnd * 255)
            Dim CardBitmap As Bitmap = New Bitmap(80, 80)
            Using CardGraphics As Graphics = Graphics.FromImage(CardBitmap)
                For PointIndex As Integer = 0 To 8
                    If ObjectModel(ObjectType(Index), PointIndex).Equals(EmptyPoint) Then Exit For '遍历到空坐标结束循环
                    CardGraphics.FillRectangle(New SolidBrush(ObjectColor(Index)), New Rectangle(ObjectModel(ObjectType(Index), PointIndex).Y * 16, ObjectModel(ObjectType(Index), PointIndex).X * 16, 15, 15))
                Next
            End Using
            ObjectLabels(Index).Image = CardBitmap
            GC.Collect() '回收内存
        Next
    End Sub

    Private Sub ObjectLabel_MouseDown(sender As Object, e As MouseEventArgs) Handles ObjectLabel0.MouseDown, ObjectLabel1.MouseDown, ObjectLabel2.MouseDown
        Dim ObjectLabel As Label = CType(sender, Label)
        ObjectLabel.Size = New Size(135, 135)
        ObjectLabel.Image = New Bitmap(ObjectLabel.Image, 135, 135)
        MousePointInLabel = e.Location
        AddHandler ObjectLabel.MouseMove, AddressOf ObjectLabel_MouseMove
    End Sub

    Private Sub ObjectLabel_MouseMove(sender As Object, e As MouseEventArgs)
        CType(sender, Label).Left = MousePosition.X - Me.Left - MousePointInLabel.X
        CType(sender, Label).Top = MousePosition.Y - Me.Top - MousePointInLabel.Y
    End Sub

    Private Sub ObjectLabel_MouseUp(sender As Object, e As MouseEventArgs) Handles ObjectLabel0.MouseUp, ObjectLabel1.MouseUp, ObjectLabel2.MouseUp
        Dim ObjectLabel As Label = CType(sender, Label)
        RemoveHandler ObjectLabel.MouseMove, AddressOf ObjectLabel_MouseMove

        ObjectLabel.Size = New Size(80, 80)
        ObjectLabel.Image = New Bitmap(ObjectLabel.Image, 80, 80)
        '恢复 ObjectLabel 的位置
        ObjectLabel.Location = ObjectLabelLocation(ObjectLabel.Tag)

        If MoveToGameAera(ObjectLabel.Tag) Then
            '拖入游戏区成功！
            IsFullInLine()

            DrawForm()
            ObjectLabel.Hide()
            ObjectCount -= 1
            If ObjectCount = 0 Then CreateNewObject()

            If IsGameOver() Then GameOver()
        End If
    End Sub

    ''' <summary>
    ''' 尝试将目标物体拖动进游戏区域
    ''' </summary>
    ''' <param name="Index">拖入的物体的标识</param>
    ''' <returns>是否拖入成功</returns>
    Private Function MoveToGameAera(ByVal Index As Integer) As Boolean
        Dim IndexX, IndexY As Integer '当前鼠标位置对应的 CardData 坐标
        Dim PointsInGameAera(0) As Point
        Dim PointInObjectModel As Point
        Dim PointInGameAera As Point
        IndexY = (MousePosition.X - PaddingSize - MousePointInLabel.X - Me.Left) \ (CardSize + MarginSize)
        IndexX = (MousePosition.Y - PaddingSize - TitleHeight - MousePointInLabel.Y - Me.Top) \ (CardSize + MarginSize)
        If IndexX < 0 OrElse IndexX > 9 OrElse IndexY < 0 OrElse IndexY > 9 Then Return False
        For PointIndex As Integer = 0 To 8
            PointInObjectModel = ObjectModel(ObjectType(Index), PointIndex)
            If PointInObjectModel.Equals(EmptyPoint) Then Exit For
            PointInGameAera.X = IndexX + PointInObjectModel.X
            PointInGameAera.Y = IndexY + PointInObjectModel.Y
            If PointInGameAera.X < 0 OrElse PointInGameAera.X > 9 OrElse PointInGameAera.Y < 0 OrElse PointInGameAera.Y > 9 Then Return False

            If CardData(PointInGameAera.X, PointInGameAera.Y) Then
                '放不下
                Return False
            Else
                PointsInGameAera(UBound(PointsInGameAera)) = New Point(PointInGameAera.X, PointInGameAera.Y)
                ReDim Preserve PointsInGameAera(UBound(PointsInGameAera) + 1)
            End If
        Next
        ReDim Preserve PointsInGameAera(UBound(PointsInGameAera) - 1)
        For PointIndex As Integer = 0 To UBound(PointsInGameAera)
            Score += 1
            CardData(PointsInGameAera(PointIndex).X, PointsInGameAera(PointIndex).Y) = True
            ColorData(PointsInGameAera(PointIndex).X, PointsInGameAera(PointIndex).Y) = ObjectColor(Index)
        Next
        '可以放下
        Return True
    End Function

    ''' <summary>
    ''' 检测是否存在整行或者整列
    ''' </summary>
    Private Sub IsFullInLine()
        Dim IndexX, IndexY As Integer
        Dim IsFulled As Boolean
        Dim FulledLine(0) As Integer
        Dim FulledColumn(0) As Integer
        '记录整行
        For IndexY = 0 To 9
            IsFulled = True
            For IndexX = 0 To 9
                If Not CardData(IndexY, IndexX) Then
                    IsFulled = False : Exit For
                End If
            Next
            If IsFulled Then
                FulledLine(UBound(FulledLine)) = IndexY
                ReDim Preserve FulledLine(UBound(FulledLine) + 1)
            End If
        Next
        ReDim Preserve FulledLine(UBound(FulledLine) - 1)
        '记录整列
        For IndexX = 0 To 9
            IsFulled = True
            For IndexY = 0 To 9
                If Not CardData(IndexY, IndexX) Then
                    IsFulled = False : Exit For
                End If
            Next
            If IsFulled Then
                FulledColumn(UBound(FulledColumn)) = IndexX
                ReDim Preserve FulledColumn(UBound(FulledColumn) + 1)
            End If
        Next
        ReDim Preserve FulledColumn(UBound(FulledColumn) - 1)
        '计算得分
        Score += ScoreList(FulledColumn.Count + FulledLine.Count)

        '清除整行或整列
        If FulledLine.Count > 0 Then
            For IndexY = 0 To UBound(FulledLine)
                For IndexX = 0 To 9
                    CardData(FulledLine(IndexY), IndexX) = False
                Next
            Next
        End If
        If FulledColumn.Count > 0 Then
            For IndexX = 0 To UBound(FulledColumn)
                For IndexY = 0 To 9
                    CardData(IndexY, FulledColumn(IndexX)) = False
                Next
            Next
        End If
    End Sub

    ''' <summary>
    ''' 检测游戏是否结束
    ''' </summary>
    ''' <returns></returns>
    Private Function IsGameOver() As Boolean
        Dim Index As Integer
        Dim PointX, PointY As Integer
        Dim Result As Boolean

        For Index = 0 To 2
            If ObjectLabels(Index).Visible Then
                For PointY = 0 To 9
                    For PointX = 0 To 9
                        Result = CanPutItIn(PointX, PointY, ObjectType(Index))
                        If Result Then Exit For
                    Next
                    If Result Then Exit For
                Next
                If Result Then Exit For
            End If
        Next
        Return Not Result
    End Function

    ''' <summary>
    ''' 检测 ObjectType 对应的物体能否放在 CardData(PointX,PointY) 里
    ''' </summary>
    ''' <param name="PointX"></param>
    ''' <param name="PointY"></param>
    ''' <param name="ObjectType"></param>
    ''' <returns>能否放置</returns>
    Private Function CanPutItIn(ByVal PointX As Integer, ByVal PointY As Integer, ByVal ObjectType As Integer) As Boolean
        Dim PointInObjectModel As Point
        Dim PointInGameAera As Point
        For PointIndex As Integer = 0 To 8
            PointInObjectModel = ObjectModel(ObjectType, PointIndex)
            If PointInObjectModel.Equals(EmptyPoint) Then Exit For
            PointInGameAera.X = PointX + PointInObjectModel.X
            PointInGameAera.Y = PointY + PointInObjectModel.Y
            If PointInGameAera.X < 0 OrElse PointInGameAera.X > 9 OrElse PointInGameAera.Y < 0 OrElse PointInGameAera.Y > 9 Then Return False
            If CardData(PointInGameAera.X, PointInGameAera.Y) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 游戏结束
    ''' </summary>
    Private Sub GameOver()
        MsgBox("得分：" & Score, 64, "游戏结束：")
        Score = 0
        ScoreLabel.Text = "0"
        ReDim CardData(9, 9)
        CreateNewObject()
        DrawForm()
    End Sub
End Class
