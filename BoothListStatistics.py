import gspread
import gspread.urls
from gspread.utils import ValueRenderOption, ValueInputOption
import gspread.utils
import gspread_formatting
from gspread_formatting import Borders, TextFormat, ColorStyle, Color
from datetime import datetime
import time

def clean_N_InList(GenreLists: list[list[str]]):
	"""
	시트에서 가져온 장르 리스트에 대하여 개행 문자를 전부 띄어쓰기를 위한 빈 칸으로 변환합니다.

	@param GenreLists 시트에서 가져온 장르 리스트, 일반적으로 `gspread.WorkSheet.get_values()` 함수를 사용합니다.
	@return 개행 문자를 모두 변환한 장르 리스트
	"""
	GenreList_temp : list[str] = []
	for i in range(0, len(GenreLists)):
		if GenreLists[i][0] != '장르' and GenreLists[i][0] != '':
			if GenreLists[i][0].find('\n') != -1:
				GenreList_temp.append([GenreLists[i][0].replace("\n", " ")])
			else:
				GenreList_temp.append([GenreLists[i][0]])

	return GenreList_temp

def distributeGenres(GenreLists: list[list[str]]):
	"""
	시트에서 가져온 장르 리스트에서 중복을 모두 제외한 장르 리스트르 만듭니다.

	@param GenreLists 시트에서 가져온 장르 리스트, 일반적으로 `gspread.WorkSheet.get_values()` 함수를 사용합니다.
	@return 중복을 모두 제외한 장르 리스트
	"""
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
	"""
	시트에서 가져온 장르 리스트에서 개수를 세어, 통계 기록을 만듭니다.

	@param GenreLists_origin 시트에서 가져온 장르 리스트, 일반적으로 `gspread.WorkSheet.get_values()` 함수를 사용합니다.
	@param GenreList 함수 `distributeGenres()`에 의해 반환된 중복 없는 장르 리스트입니다.

	@return {장르 : 개수}로 이루어진 `Dictionary`
	"""
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
	
# 장르 자료를 가져올 부스 목록 시트의 ID 및 하위 워크 시트의 인덱스 넘버 (0부터 시작)
spreadsheetId = "1TmZxEkJW17d0I1MmfNyzIIxjh1n_en1DKrwsbk2OzjM"
sheetNumber = 0

# 통계 자료를 넣을 시트의 ID 및 하위 워크 시트의 인덱스 넘버
statistics_spreadsheetId = "1qmIvmDX9GdS8yoMlIaVTGAwphpTU7xwfq-uLrubfDTk"
statictics_sheetNumber = 0

statictics_sheetUpdatetime_al = 'C3'

# 통계 워크 시트에서 자료를 넣기 시작하는 지점 (1부터 시작)
statictics_sheetStartIndex = 6

client_ = gspread.service_account()
sh = client_.open_by_key(spreadsheetId)
sheet = sh.get_worksheet(sheetNumber)

sh_genre = client_.open_by_key(statistics_spreadsheetId)
sheet_genre = sh_genre.get_worksheet(statictics_sheetNumber)

GenreDatas = sheet.get_values('D:D', value_render_option=ValueRenderOption.formatted)
Genre_List = distributeGenres(GenreDatas)
Genre_Dic = countGenrefromList(GenreLists_origin=GenreDatas, GenreList=Genre_List)

print(f"Distribute result : {Genre_List}")
print("")
print(f"Distributeed result in all booths : {Genre_Dic}")
print("")

sorted_result = dict(sorted(Genre_Dic.items(), key = lambda item: item[1], reverse=True))
print(f"sorted result : {sorted_result}")

fmt = gspread_formatting.CellFormat(
	borders=Borders(
		top=gspread_formatting.Border("SOLID"),
		bottom=gspread_formatting.Border("SOLID"),
		left=gspread_formatting.Border("SOLID"),
		right=gspread_formatting.Border("SOLID")
	),
	horizontalAlignment='CENTER',
	verticalAlignment='MIDDLE',
	backgroundColorStyle=ColorStyle(rgbColor=Color(red=1, green=1, blue=1)),
	textFormat=TextFormat(
		bold='false'
	),
)

genre_datas_already = sheet_genre.get('D:D', major_dimension=gspread.utils.Dimension.cols)
sheet_genre.delete_rows(statictics_sheetStartIndex, len(genre_datas_already[0]))
grade_Index = 1					# 순위 인덱스
duplicate_index = 1				# 장르 개수가 중복되는 경우, 1씩 증가하는 중복 인덱스
index = 1						# 인덱스
former_value = 0				# 이전 인덱스의 장르 개수 (개수가 같은지 아닌지를 비교하는데 사용됨)

updatetime = datetime.now()
sheet_genre.update_acell(f'{statictics_sheetUpdatetime_al}',
					 f'{updatetime.year}. {updatetime.month}. {updatetime.day} {updatetime.hour}:{str(updatetime.minute).zfill(2)}:{str(updatetime.second).zfill(2)}')

for key in sorted_result:
	if index != 1:												# 인덱스가 1인 경우, 이전 데이터가 없으므로 인덱스가 1인 경우를 제외함
		if sorted_result[key] == former_value:					# 이전 장르 개수와 동일하면, 중복 인덱스 1 증가
			duplicate_index += 1
		else:													# 이전 장르 개수와 다르면
			if duplicate_index != 1:							# 이 인덱스의 순위는 원래 순위 + 중복 인덱스로 계산
				grade_Index = grade_Index + duplicate_index
				duplicate_index = 1
			else:												# 아니면 순위 인덱스 1 증가
				grade_Index += 1

	NewData = [str(grade_Index), key, f"{sorted_result[key]}개"]
	sheet_genre.append_row(NewData, value_input_option= ValueInputOption.user_entered)
	gspread_formatting.set_row_height(sheet_genre, str(statictics_sheetStartIndex + index - 1), 30)
	gspread_formatting.format_cell_range(sheet_genre, f'B{statictics_sheetStartIndex + index - 1}:D{statictics_sheetStartIndex + index - 1}', fmt)
	former_value = sorted_result[key]
	index += 1

	print(f"장르 {key} 작업 중..... [{index} / {len(sorted_result)}]")

	time.sleep(3.2)												# 1분당 API 요청 한계로 인한 오류가 발생하지 않도록 준 딜레이
