import gspread
import gspread.urls
from gspread.utils import ValueRenderOption, ValueInputOption
import gspread.utils
import gspread_formatting
from gspread_formatting import Borders
import time

def clean_N_InList(GenreLists: list[list[str]]):
	GenreList_temp : list[str] = []
	for i in range(0, len(GenreLists)):
		if GenreLists[i][0] != '장르' and GenreLists[i][0] != '':
			if GenreLists[i][0].find('\n') != -1:
				GenreList_temp.append([GenreLists[i][0].replace("\n", " ")])
			else:
				GenreList_temp.append([GenreLists[i][0]])

	return GenreList_temp

def ditributeGenres(GenreLists: list[list[str]]):
	#print(f"GenreList : {GenreLists}")
	#print("")
	GenreList = []
	GenreList_temp = clean_N_InList(GenreLists= GenreLists)

	for j in range (0, len(GenreList_temp)):
		GenreSubList = GenreList_temp[j][0].split(", ")
		for k in range(0, len(GenreSubList)):
			if "Vtuber" in GenreSubList[k]:
				if "Vtuber" not in GenreList:
					GenreList.append("Vtuber")
			elif "(" in GenreSubList[k] or ")" in GenreSubList[k]:
				continue
			elif GenreSubList[k] not in GenreList:
				GenreList.append(GenreSubList[k])
			
	return GenreList

def countGenrefromList(GenreLists_origin: list[list[str]], GenreList: list[str]):
	GenreList_Count_Dic = {Genre : 0 for Genre in GenreList}
	GenreList_temp = clean_N_InList(GenreLists= GenreLists_origin)

	for j in range (0, len(GenreList_temp)):
		GenreSubList = GenreList_temp[j][0].split(", ")
		for k in range(0, len(GenreSubList)):
			if "Vtuber" in GenreSubList[k]:
				GenreList_Count_Dic['Vtuber'] += 1
			elif "(" in GenreSubList[k] or ")" in GenreSubList[k]:
				continue
			elif GenreSubList[k] in GenreList:
				GenreList_Count_Dic[GenreSubList[k]] += 1

	return GenreList_Count_Dic
	

spreadsheetId = "1TmZxEkJW17d0I1MmfNyzIIxjh1n_en1DKrwsbk2OzjM"
sheetNumber = 0

statistics_spreadsheetId = "1qmIvmDX9GdS8yoMlIaVTGAwphpTU7xwfq-uLrubfDTk"
statictics_sheetNumber = 0

statictics_sheetStartIndex = 5

client_ = gspread.service_account()
sh = client_.open_by_key(spreadsheetId)
sheet = sh.get_worksheet(sheetNumber)

sh_genre = client_.open_by_key(statistics_spreadsheetId)
sheet_genre = sh_genre.get_worksheet(statictics_sheetNumber)

GenreDatas = sheet.get_values('D:D', value_render_option=ValueRenderOption.formatted)
Genre_List = ditributeGenres(GenreDatas)
Genre_Dic = countGenrefromList(GenreLists_origin=GenreDatas, GenreList=Genre_List)

print(f"Distribute result : {Genre_List}")
print("")
print(f"Distributeed result in all booths : {Genre_Dic}")
print("")

sorted_result = dict(sorted(Genre_Dic.items(), key = lambda item: item[1], reverse=True))
print(f"sorted result : {sorted_result}")

fmt = gspread_formatting.CellFormat(0
	borders=Borders(
		top=gspread_formatting.Border("SOLID"),
		bottom=gspread_formatting.Border("SOLID"),
		left=gspread_formatting.Border("SOLID"),
		right=gspread_formatting.Border("SOLID")
	),
	horizontalAlignment='CENTER',
	verticalAlignment='MIDDLE',
)

genre_datas_already = sheet_genre.get('D:D', major_dimension=gspread.utils.Dimension.cols)
sheet_genre.delete_rows(statictics_sheetStartIndex, len(genre_datas_already[0]))
grade_Index = 1
duplicate_index = 1
index = 1
former_value = 0
for key in sorted_result:
	if index != 1:
		if sorted_result[key] == former_value:
			duplicate_index += 1
		else:
			if duplicate_index != 1:
				grade_Index = grade_Index + duplicate_index
				duplicate_index = 1
			else:
				grade_Index += 1

	NewData = [str(grade_Index), key, f"{sorted_result[key]}개"]
	sheet_genre.append_row(NewData, value_input_option= ValueInputOption.user_entered)
	gspread_formatting.set_row_height(sheet_genre, str(statictics_sheetStartIndex + index - 1), 30)
	gspread_formatting.format_cell_range(sheet_genre, f'B{statictics_sheetStartIndex + index - 1}:D{statictics_sheetStartIndex + index - 1}', fmt)
	former_value = sorted_result[key]
	index += 1

	time.sleep(3.2)
