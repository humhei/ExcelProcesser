module Type
type Cell={ 
     Value : string
     RowNum : int
     ColNum : int
}
type WorkSheet={ Cells : Cell list}
type WorkBook={ WorkSheets : WorkSheet list }