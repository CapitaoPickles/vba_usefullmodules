Attribute VB_Name = "Módulo5"
Sub Fourier_Transform(rng As Range)
' -----------------------------------------------------------Setup e constantes-------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------
    Dim N, k, m, ko, i As Long
    Dim cell As Range
    Dim ft_sin(), ft_cos(), aux_sin, aux_cos, modulo(), fase() As Double
    Const pi = 3.14159265358979
    N = count_rows(rng)
    ReDim ft_sin(N), ft_cos(N), modulo(N), fase(N)
'-------------------------------------------------------------------------------------------------------------------------------------------------

    k = 0
    ko = 0  'Offset. Pode ser transladado
    m = 0
    aux_sin = 0
    aux_cos = 0
    While k <= N - (ko + 1)
      For Each cell In rng
        aux_cos = aux_cos + cell.value * Cos(2 * pi * k * m / N)  'Parte real da transformada
        aux_sin = aux_sin - cell.value * Sin(2 * pi * k * m / N)  'Parte complexa da transformada
        m = m + 1
      Next cell
      m = 0
      ft_sin(k) = aux_sin
      ft_cos(k) = aux_cos
      modulo(k) = Sqr((ft_sin(k)) ^ 2 + (ft_cos(k)) ^ 2)  'Cálculo do módulo da transformada de Fourier
      fase(k) = ft_sin(k) / ft_cos(k)   'Cálculo da fase da transformada de Fourier
      aux_sin = 0
      aux_cos = 0
      k = k + 1
    Wend
    
End Sub
