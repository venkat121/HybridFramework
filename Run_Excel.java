package datadriven_framework;

public class Run_Excel 
{
	
	public static void main(String args[])
	{
		
		Excel.Get_ExcelConnection("InputData.xlsx", "config");
		
		String url=Excel.Get_CellData(1, 0);
		System.out.println(url);
		
		Excel.WriteCellData(1, 4, "TestPass");
		
		Excel.Create_OPfile("OP.xlsx");
		
		
	}

}
