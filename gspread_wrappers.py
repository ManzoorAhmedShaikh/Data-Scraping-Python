from gspread_formatting import *

class GspreadWrapper:

    def UpdateCells(self,worksheet, cell_range,values):
        """
        Update Cells in 1 request
        :param worksheet: Worksheet where to update cells
        :param cell_range: Range of cells in A1 notation "A1:Ax"
        :param values: A 2d List that needs to be inserted in sheet
        (e.g for 1 Row [["A","B","C"]]) and (e.g for 1 column [["A"],["B"],["C"]])
        :return:
        """
        try:
            worksheet.update(str(cell_range),values)
        except:
            print("Invalid Values")

    def BatchUpdateCells(self,worksheet,ranges):
        """
        Update multiple range of cells in 1 request
        :param worksheet: Worksheet where to update cells
        :param ranges: An iterable list contains a set of ranges and their values
        e.g [("B2:B4",["A","B","C"]),("D4:D5",[2,4]),...]
        :return:
        """

        final_format = []
        for x in ranges:
            final_format.append({'range': x[0], 'values': x[1]})

        worksheet.batch_update(final_format)

    def CreateFormattingStyle(self,bg=(1,1,1),fg=(0,0,0),fsize=12,border=False):
        """
        Create the Formatting Style
        :param worksheet: Worksheet object where to apply the formatting
        :param cell_range: Range of cells in A1 notation "A1:Ax"
        :param bg: Cell background color in a tuple of rgb (default: (1,1,1) white)
        :param fg: Font color in a tuple of rgb (default: (0,0,0) black)
        :param fsize: Font size in integer (default: 12)
        :param border: If True, border will place all around the cell (Default: False)
        :return: CellFormat object that can be used for applying the formatting
        """
        try:
            if border:
                return CellFormat(backgroundColor=Color(red=bg[0],green=bg[1],blue=bg[2]),textFormat=TextFormat(
                foregroundColor=Color(red=fg[0],green=fg[1],blue=fg[2]),fontSize=fsize),borders=Borders(top=Border(style='SOLID',width=2),
                right=Border(style='SOLID',width=2),bottom=Border(style='SOLID',width=2),
                left=Border(style='SOLID',width=2)))

            else:
                return CellFormat(backgroundColor=Color(red = bg[0],green=bg[1],blue=bg[2]),textFormat=TextFormat(
                    foregroundColor=Color(red=fg[0],green=fg[1],blue=fg[2]),fontSize=fsize))
        except:
            print("Invalid Arguments")

    def ApplySingleFormatting(self,worksheet,cell_range,format):
        """
        Apply Formatting in 1 Request
        :param worksheet: Worksheet object where to apply the formatting
        :param cell_range: Range of cells in A1 notation "A1:Ax"
        :param format: format style that needs to be applied
        :return:
        """

        try:
            format_cell_range(worksheet,cell_range,format)

        except:
            print("Invalid Argument")

    def ApplyingMultiFormatting(self,worksheet,range_format):
        """
        Apply Formatting on multiple cells in 1 Request
        :param worksheet: Worksheet object where to apply the formatting
        :param range_format: An iterable in a pair of ranges and their format
        e.g [("A1:A3",format1),("B1:B3",format2),...]
        :return:
        """

        try:
            format_cell_ranges(worksheet,range_format)

        except:
            print("Invalid Arguments")
