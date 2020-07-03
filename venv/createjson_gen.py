import xlrd
import copy
import json
import unicodedata
import datetime
import sys
# these components are for file operations
import os
import glob

# XCDE Map Build Process
# This process parses an excel file and produces a json map for consumption by XCDE

acceptable_filenames = [
	"PDXC_CIs_Attr_Rel_Mapping.xlsx",
	"PDXC_ITAM_CIs_Attr_Rel_Mapping.xlsx",
	"CIs_Attr_Rel_Mapping_ES_UCMDB_ESL.xls"
]
file_directory = 'processqueue'
output_directory = 'processoutput'
print(
	"File to processed must be in the " + file_directory + " directory.\nSelect one of the below file names for processing: ")
file_number = 1
for this_file in acceptable_filenames:
	print(str(file_number) + ") " + this_file)
	file_number += 1
print(str(file_number) + ") Other")
print(str(99) + ") Quit")

file_location = ''
file_name = ''
while file_name == '':
	selected_number = 0
	try:
		selected_number = int(input("Enter numeric selection: "))
		if selected_number == 99:
			print('Exiting program')
			quit(200)
		if selected_number < 1 or selected_number > file_number:
			print("Invalid input. Must be between 1 and " + str(file_number))
		elif selected_number == file_number:
			file_name = input("Enter file name: ")
		else:
			file_name = acceptable_filenames[selected_number - 1]
	except ValueError:
		print("Invalid input. Must be an integer.")

	if file_name != '':
		file_location = file_directory + "/" + file_name
		print("Opening " + file_location)
		try:
			workbook = xlrd.open_workbook(file_location)
		except FileNotFoundError:
			print("File not found.  Ensure " + file_name + " is in the " + file_directory + " directory.")
			file_location = ''
			file_name = ''
		if file_name != '':
			# Verify that it's ok to clear the process output directory
			existing_output_files = glob.glob(output_directory + "/*.*")
			output_full_path = os.getcwd()+"/"+output_directory
			if existing_output_files != []:
				#print(existing_output_files)
				print("Do you wish to delete all existing files in the "+output_full_path+" directory?")
				confirmation = input('(y/n)>: ')
				if confirmation.lower() == 'y':
					for efile in os.listdir(output_full_path):
						#print(efile)
						os.remove(output_full_path+"/"+efile)

					print("FYI - All files in output directory have been deleted.")
				else:
					print("FYI - No existing files have been deleted. Matching files produced by this process have been overwritten.")
			else:
				print("FYI - Nothing to delete in output directory.")

#print("Temp Quit")
#exit()
mapping_json = []

now = datetime.datetime.now()

log = open(output_directory + '/logfile_' + str(now.month) + '_' + str(now.day) + '_' + str(now.year) + '.log', 'w')

log.write("[INFO] - File location: " + file_location + "\n")
# print("[INFO] - File location: "+ file_location)
# workbook = xlrd.open_workbook(file_location)

# get configuration
mapConfigSheet = workbook.sheet_by_name('Mapping_Configurations')
log.write("[INFO] - Getting mapping configuration \n")
# print("[INFO] - Getting mapping configuration \n")

# Values found in column: C
ColumnValues = 2
# Attributes sheet
multipleAttrSheet = {}
AttributesSheet = str(mapConfigSheet.cell(1, ColumnValues).value.replace('\n', ' ').replace('\r', '')).strip()
if ";" in AttributesSheet:
	for index, item in enumerate(AttributesSheet.split(";")):
		multipleAttrSheet["AttrSheet" + str(index)] = item
	log.write("[INFO] - Attributes sheets: " + json.dumps(multipleAttrSheet, indent=4, sort_keys=False) + "\n")
# print("[INFO] - Attributes sheets: " + json.dumps(multipleAttrSheet,indent=4, sort_keys=False) + "\n")
else:
	log.write("[INFO] - Attributes sheet: " + AttributesSheet + "\n")
# print("[INFO] - Attributes sheet: " + AttributesSheet + "\n")

# CI Types sheet
multipleCITSheet = {}
CITypesSheet = str(mapConfigSheet.cell(2, ColumnValues).value.replace('\n', ' ').replace('\r', '')).strip()
if ";" in CITypesSheet:
	for index, item in enumerate(CITypesSheet.split(";")):
		multipleCITSheet["CISheet" + str(index)] = item
	log.write("[INFO] - CI Types sheets: " + json.dumps(multipleCITSheet, indent=4, sort_keys=False) + "\n")
# print("[INFO] - CI Types sheets: " + json.dumps(multipleCITSheet,indent=4, sort_keys=False) + "\n")
else:
	log.write("[INFO] - CI Types sheet: " + CITypesSheet + "\n")
# print("[INFO] - CI Types sheet: " + CITypesSheet + "\n")

# Relationships sheet
multipleRelSheet = {}
RelationshipsSheet = str(mapConfigSheet.cell(3, ColumnValues).value.replace('\n', ' ').replace('\r', '')).strip()
if ";" in RelationshipsSheet:
	for index, item in enumerate(RelationshipsSheet.split(";")):
		multipleRelSheet["RelSheet" + str(index)] = item
	log.write("[INFO] - Relationships sheets: " + json.dumps(multipleRelSheet, indent=4, sort_keys=False) + "\n")
# print("[INFO] - Relationships sheets: " + json.dumps(multipleRelSheet,indent=4, sort_keys=False) + "\n")
else:
	log.write("[INFO] - Relationships sheet: " + RelationshipsSheet + "\n")
# print("[INFO] - Relationships sheet: " + RelationshipsSheet + "\n")

# Enum sheet
multipleEnumSheet = {}
EnumSheet = str(mapConfigSheet.cell(4, ColumnValues).value.replace('\n', ' ').replace('\r', '')).strip()
if ";" in EnumSheet:
	for index, item in enumerate(EnumSheet.split(";")):
		multipleEnumSheet["EnumSheet" + str(index)] = item
	log.write("[INFO] - Enum sheets: " + json.dumps(multipleEnumSheet, indent=4, sort_keys=False) + "\n")
# print("[INFO] - Enum sheets: " + json.dumps(multipleEnumSheet,indent=4, sort_keys=False) + "\n")
else:
	log.write("[INFO] - Enum sheet: " + EnumSheet + "\n")
# print("[INFO] - Enum sheet: " + EnumSheet + "\n")

# get filters
FiltersColumn = 3

FiltersBySheet = {}

AttributesFilters = []
cellAttrValues = str(mapConfigSheet.cell(1, FiltersColumn).value.replace('\n', ' ').replace('\r', '')).strip().split(
	";")
# print(json.dumps(cellAttrValues,indent=4, sort_keys=False))
for filterV in cellAttrValues:
	if len(filterV) > 1:
		colvals = filterV.split(",")
		AttributesFilters.append({"Column": str.replace(colvals[0], "Column: ", " ").strip(),
								  "value": str.replace(colvals[1], "value: ", " ").strip()})
log.write("[INFO] - Attributes Filters: \n")
log.write(json.dumps(AttributesFilters, indent=4, sort_keys=False))
##print(json.dumps(AttributesFilters,indent=4, sort_keys=False))

FiltersBySheet[AttributesSheet] = AttributesFilters

CITypesFilters = []
cellAttrValues = str(mapConfigSheet.cell(2, FiltersColumn).value.replace('\n', ' ').replace('\r', '')).strip().split(
	";")
# print(json.dumps(cellAttrValues,indent=4, sort_keys=False))
for filterV in cellAttrValues:
	if len(filterV) > 1:
		colvals = filterV.split(",")
		CITypesFilters.append({"Column": str.replace(colvals[0], "Column: ", " ").strip(),
							   "value": str.replace(colvals[1], "value: ", " ").strip()})
log.write("[INFO] - CI Types Filters: \n")
log.write(json.dumps(CITypesFilters, indent=4, sort_keys=False))
##print(json.dumps(CITypesFilters,indent=4, sort_keys=False))

FiltersBySheet[CITypesSheet] = CITypesFilters

RelFilters = []
cellAttrValues = str(mapConfigSheet.cell(3, FiltersColumn).value.replace('\n', ' ').replace('\r', '')).strip().split(
	";")
##print(json.dumps(cellAttrValues,indent=4, sort_keys=False))
for filterV in cellAttrValues:
	if len(filterV) > 1:
		colvals = filterV.split(",")
		RelFilters.append({"Column": str.replace(colvals[0], "Column: ", " ").strip(),
						   "value": str.replace(colvals[1], "value: ", " ").strip()})
log.write("[INFO] - Relationships Filters: \n")
log.write(json.dumps(RelFilters, indent=4, sort_keys=False))
##print(json.dumps(RelFilters,indent=4, sort_keys=False))

FiltersBySheet["Relationships"] = RelFilters

ENUMFilters = []
cellAttrValues = str(mapConfigSheet.cell(4, FiltersColumn).value.replace('\n', ' ').replace('\r', '')).strip().split(
	";")
##print(json.dumps(cellAttrValues,indent=4, sort_keys=False))
for filterV in cellAttrValues:
	if len(filterV) > 1:
		colvals = filterV.split(",")
		ENUMFilters.append({"Column": str.replace(colvals[0], "Column: ", " ").strip(),
							"value": str.replace(colvals[1], "value: ", " ").strip()})
log.write("[INFO] - ENUM Filters: \n")
log.write(json.dumps(ENUMFilters, indent=4, sort_keys=False))


##print(json.dumps(ENUMFilters,indent=4, sort_keys=False))

def getColNumberByName(name, sheetName):
	colnumber = 0
	sheet = workbook.sheet_by_name(sheetName.strip())

	for colnum in range(sheet.ncols):
		name = name.replace('\n', ' ').replace('\r', '').strip()
		columnHeader = sheet.cell(0, colnum).value.replace('\n', ' ').replace('\r', '')
		if name in columnHeader:
			colnumber = colnum

	return colnumber


def getCIsMatchingValues(sheetName, ColumnValue, filterValue, filterColumn):
	# print(sheetName + " " + ColumnValue + " " + filterValue  + " " + filterColumn)
	matchingValues = []
	sheet = workbook.sheet_by_name(sheetName.strip())
	for rownum in range(sheet.nrows):
		columnValue = str(sheet.cell(rownum, getColNumberByName(ColumnValue.strip(), sheetName.strip())).value).strip()
		columnFilterValue = str(
			sheet.cell(rownum, getColNumberByName(filterColumn.strip(), sheetName.strip())).value).strip()
		filterValues = []
		if "[" in filterValue:
			filterValues = filterValue.replace("[", " ").replace("]", " ").strip().split(" ")

		if columnFilterValue == filterValue.strip() or columnFilterValue in filterValues:
			if ">" in columnValue:
				displayValue = columnValue.split(">")
				columnValue = displayValue[len(displayValue) - 1]
			if "\n" in columnValue:
				cis = columnValue.split("\n")
				for c in cis:
					matchingValues.append(c.strip())
			else:
				if columnValue != "":
					matchingValues.append(columnValue)
	return matchingValues


FiltersBySheet[EnumSheet] = ENUMFilters

# Filters by mapping direction
mappingDirections = []
FiltersByMappingDirection = []
FiltersByMappingDirectionFlag = False
UsecasesAndConnectionBySource = []
connectionsByDirectionAndIntegration = []
# get columns to get values per sheet and mapping direction
columnsToGetValuesBySheetNMapDireccions = []

for rownum in range(mapConfigSheet.nrows):
	mappingDirection = {}
	columnsToGetValuesBySheetNMapDireccion = {}
	FilterByMappingDirection = {}
	filtersCol = str(mapConfigSheet.cell(rownum, FiltersColumn).value.replace('\n', ' ').replace('\r', '')).strip()
	if filtersCol == "Filter per target":
		FiltersByMappingDirectionFlag = True
	else:
		if FiltersByMappingDirectionFlag == True:
			integrationName = str(mapConfigSheet.cell(rownum, 0).value).strip()
			sourceSysNameDetails = str(mapConfigSheet.cell(rownum, 1).value).strip().split(";")
			targetSysName = str(mapConfigSheet.cell(rownum, 2).value).strip()
			usecaseDetails = str(mapConfigSheet.cell(rownum, 6).value).strip().split(";")

			# get usecase and connection details
			targetsDet = []
			targetDetailsValues = str(mapConfigSheet.cell(rownum, 5).value).strip().split(";")
			for td in targetDetailsValues:
				allDetails = td.replace("details=", " ").split(",")
				targetObj = {}
				for d in allDetails:
					# connection["target"] = targetSysName
					if len(d) > 1:
						# print(con)
						details = d.split("=")
						# print(details)
						if '{' in details[1] and '}' in details[1]:
							subObj = details[1].split('{')[0]
							findItem = False
							if targetObj.get(details[0].replace("\n", " ").strip()) != None:
								findItem = targetObj.get(details[0].replace("\n", " ").strip())
								findItem[subObj] = {}
							else:
								targetObj[details[0].replace("\n", " ").strip()] = {}
								targetObj[details[0].replace("\n", " ").strip()][subObj] = {}

							detailsIn = details[1].strip().split('{')[1].replace('}', ' ').strip().split("|")

							# print(details[0])
							for innerItem in detailsIn:
								moreDetails = innerItem.split(":")
								if moreDetails[0] == "matchingValues":
									if '[' in moreDetails[1]:
										# print(moreDetails[1])
										targetObj[details[0].replace("\n", " ").strip()][subObj][
											moreDetails[0].replace("\n", " ").strip()] = [
											moreDetails[1].replace('[', ' ').replace(']', ' ').strip()]
									else:
										getMatchingValues = moreDetails[1].split("*")
										targetObj[details[0].replace("\n", " ").strip()][subObj][
											moreDetails[0].replace("\n", " ").strip()] = getCIsMatchingValues(
											getMatchingValues[0], getMatchingValues[1], getMatchingValues[2],
											getMatchingValues[3])
								else:
									if innerItem != '':
										valueMoreDetails = moreDetails[1].replace("\n", " ").strip()
										if valueMoreDetails == "false":
											valueMoreDetails = False
										if valueMoreDetails == "true":
											valueMoreDetails = True
										targetObj[details[0].replace("\n", " ").strip()][subObj][
											moreDetails[0].replace("\n", " ").strip()] = valueMoreDetails
						elif "|" in details[1]:
							detailsIn = details[1].strip().split("|")
							targetObj[details[0].replace("\n", " ").strip()] = {}

							# print(detailsIn)
							for innerItem in detailsIn:
								moreDetails = innerItem.split(":")
								if moreDetails[0] == "matchingValues":
									getMatchingValues = moreDetails[1].split("*")
									targetObj[details[0].replace("\n", " ").strip()][
										moreDetails[0].replace("\n", " ").strip()] = getCIsMatchingValues(
										getMatchingValues[0], getMatchingValues[1], getMatchingValues[2],
										getMatchingValues[3])
								else:
									valueMoreDetails = moreDetails[1].replace("\n", " ").strip()
									if valueMoreDetails == "false":
										valueMoreDetails = False
									if valueMoreDetails == "true":
										valueMoreDetails = True
									targetObj[details[0].replace("\n", " ").strip()][
										moreDetails[0].replace("\n", " ").strip()] = valueMoreDetails
						else:
							# print(details)
							if '[' in details[1] and ':' in details[1]:
								print(details[1])
								subDetails = details[1].split(':')
								targetObj[details[0].replace("\n", " ").strip()] = {}
								targetObj[details[0].replace("\n", " ").strip()][subDetails[0]] = {}
								if '[' in subDetails[1]:
									subArray = subDetails[1].replace('[', ' ').replace(']', ' ').strip().split('*')
									targetObj[details[0].replace("\n", " ").strip()][subDetails[0]] = subArray
								else:
									targetObj[details[0].replace("\n", " ").strip()][subDetails[0]] = subDetails[1]
							else:
								targetObj[details[0].replace("\n", " ").strip()] = details[1].replace("\n",
																									  " ").replace(":",
																												   "=").strip()
				targetsDet.append(targetObj)
			# print(connection)
			# print(json.dumps(targetsDet,indent=4, sort_keys=False))

			sourceDetails = {}
			for sourceD in sourceSysNameDetails:
				s = sourceD.replace("\n", " ").strip().split("=")
				if len(s) > 1:
					if "," in s[1]:
						sourceDetails[s[0].replace("\n", " ").strip()] = {}
						details = s[1].replace("\n", " ").strip().split(",")
						for d in details:
							det = d.replace("\n", " ").strip().split(":")
							sourceDetails[s[0].replace("\n", " ").strip()][det[0].replace("\n", " ").strip()] = det[
								1].replace("\n", " ").strip()
					else:
						sourceDetails[s[0]] = s[1]

			for idx, usecase in enumerate(usecaseDetails):
				usecaseversion = usecase.split(",")
				usecaseVersionDetails = {}
				usecaseVersionDetails["source"] = sourceDetails["source"]
				usecaseVersionDetails["integration"] = integrationName

				usecaseVersionDetails["usecases"] = []
				usecaseObj = {}
				for uv in usecaseversion:
					uvDetails = uv.split("=")
					usecaseObj[uvDetails[0].replace("\n", " ").strip()] = uvDetails[1].replace("\n", " ").strip()
				if len(targetsDet) > 1:
					usecaseObj["targetDetails"] = targetsDet[idx]
				else:
					usecaseObj["targetDetails"] = targetsDet[0]
				usecaseObj["target"] = targetSysName
				findSourceinUsecasesAndConnectionBySource = next((item for item in UsecasesAndConnectionBySource if
																  item["source"] == sourceDetails["source"] and item[
																	  "integration"] == integrationName), False)
				if findSourceinUsecasesAndConnectionBySource == False:
					usecaseVersionDetails["usecases"].append(usecaseObj)
					UsecasesAndConnectionBySource.append(usecaseVersionDetails)
				else:
					findSourceinUsecasesAndConnectionBySource["usecases"].append(usecaseObj)

			sourceSysName = sourceDetails["source"]
			# create mapping entries for each source
			mapping_json.append({
				"source": sourceDetails["source"] + integrationName,
				"mapping": {
					"details": sourceDetails["details"],
					"solution": sourceDetails["solution"],
					"source": sourceDetails["source"],
					"targetDetails": {},
					"usecase": "",
					"version": "",
					"map": []}})

			columnsValues = str(mapConfigSheet.cell(rownum, 4).value).strip()

			columnsToGetValuesBySheetNMapDireccion["direction"] = sourceSysName + ":" + targetSysName
			columnsToGetValuesBySheetNMapDireccion["columns"] = []
			columns = columnsValues.split(";")
			columnsBySheetObj = {}
			for c in columns:
				columnsBySheetObj = {}
				columnsBySheet = c.strip().split(",")
				if len(columnsBySheet) > 1:
					for col in columnsBySheet:
						values = col.strip().split(":")
						columnsBySheetObj[values[0].strip()] = values[1].strip()
					columnsToGetValuesBySheetNMapDireccion["columns"].append(columnsBySheetObj)
			columnsToGetValuesBySheetNMapDireccions.append(columnsToGetValuesBySheetNMapDireccion)
			FilterByMappingDirection["direction"] = sourceSysName + ":" + targetSysName
			FilterByMappingDirection["filters"] = []
			filters = filtersCol.split(";")
			filterBySheetObj = {}
			for f in filters:
				filterBySheetObj = {}
				filterbySheet = f.strip().split(",")
				if len(filterbySheet) > 1:
					for v in filterbySheet:
						values = v.strip().split(":")
						filterBySheetObj[values[0].strip()] = values[1].strip()
					FilterByMappingDirection["filters"].append(filterBySheetObj)
			FiltersByMappingDirection.append(FilterByMappingDirection)

			##print("filters by mapping direction")
			##print(json.dumps(FiltersByMappingDirection,indent=4, sort_keys=False))
			##print("columns by mapping direction")
			##print(json.dumps(columnsToGetValuesBySheetNMapDireccions,indent=4, sort_keys=False))
			findSourceSys = next((item for item in mappingDirections if
								  item["source"] == sourceSysName and item["integration"] == integrationName), False)
			if findSourceSys == False:
				mappingDirection["integration"] = integrationName
				mappingDirection["source"] = sourceSysName
				mappingDirection["targets"] = []
				mappingDirection["targets"].append(targetSysName)
			else:
				if targetSysName not in findSourceSys["targets"] and findSourceSys["integration"] == integrationName:
					findSourceSys["targets"].append(targetSysName)
			if bool(mappingDirection):
				mappingDirections.append(mappingDirection)

			# Get connection for each environment

			findConnectionsByDirectionAndIntegration = next((item for item in connectionsByDirectionAndIntegration if
															 item.get(
																 sourceSysName + ":" + targetSysName + "-" + integrationName)),
															False)
			if findConnectionsByDirectionAndIntegration == False:
				con = {}
				direction = sourceSysName + ":" + targetSysName + "-" + integrationName
				con[direction] = {}

				con[direction]["eit"] = {} if str(mapConfigSheet.cell(rownum, 7).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 7).value).strip())
				con[direction]["sandbox"] = {} if str(mapConfigSheet.cell(rownum, 8).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 8).value).strip())
				con[direction]["dev"] = {} if str(mapConfigSheet.cell(rownum, 9).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 9).value).strip())
				con[direction]["dev2"] = {} if str(mapConfigSheet.cell(rownum, 10).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 10).value).strip())
				con[direction]["test"] = {} if str(mapConfigSheet.cell(rownum, 11).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 11).value).strip())
				con[direction]["stage"] = {} if str(mapConfigSheet.cell(rownum, 12).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 12).value).strip())
				con[direction]["prod"] = {} if str(mapConfigSheet.cell(rownum, 13).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 13).value).strip())
				con[direction]["elabs"] = {} if str(mapConfigSheet.cell(rownum, 14).value) == "" else json.loads(
					str(mapConfigSheet.cell(rownum, 14).value).strip())
				# print(con)
				connectionsByDirectionAndIntegration.append(con)

# print(json.dumps(connectionsByDirectionAndIntegration,indent=4, sort_keys=False))
# print("mapping directions")
# print(json.dumps(mappingDirections,indent=4, sort_keys=False))
# print(json.dumps(UsecasesAndConnectionBySource,indent=4, sort_keys=False))1
with open(output_directory + "/usecases_integrations.txt", 'w') as outfile:
	json.dump(UsecasesAndConnectionBySource, outfile, indent=4, sort_keys=True, separators=(',', ': '))


def processRelationshipsSheet(sheetName, sheetID, sourceSys, integration, mappingRootObj, columnsObj, targets,
							  excludeTarget):
	# print("processing rel method")
	sheet = workbook.sheet_by_name(sheetName.strip())
	mapBySource = {}
	mapObj = {}
	mapObj["map"] = []
	findMappingSourceSysRoot = next((item for item in mappingRootObj if item["source"] == sourceSys + integration),
									False)

	# Exception UCMDB
	ciParentList = []
	sheetCIs = workbook.sheet_by_name("CI Types")
	for rownum in range(sheetCIs.nrows):
		className = str(sheetCIs.cell(rownum, getColNumberByName("UCMDB Class", "CI Types")).value).strip()
		parenhoodVal = str(sheetCIs.cell(rownum, getColNumberByName("UCMDB Display Name", "CI Types")).value).strip()
		ciParentList.append({"ci": className, "root": parenhoodVal})
	# parentsList.append({"ci": className,"root": parenhoodVal})
	hasChilds = []
	for root in ciParentList:
		ci = root["ci"]
		for ciItem in ciParentList:
			if root["root"] in ciItem["root"]:
				if ciItem["ci"] != ci:
					findCIinhasChilds = next((item for item in hasChilds if item == ci), False)
					if findCIinhasChilds == False:
						hasChilds.append(ci)

	for rownum in range(sheet.nrows):
		# genaral filter by sheet
		for target in targets:
			if target != excludeTarget and sourceSys != excludeTarget:
				mappingItem = {}
				targetItem = {}
				# filter by target
				passBySheet = False
				for filterBySheet in FiltersBySheet[sheetID]:
					cellVal = str(
						sheet.cell(rownum, getColNumberByName(filterBySheet["Column"], sheetName)).value).strip()
					if cellVal == filterBySheet["value"]:
						passBySheet = True

				# filter by direction
				passByTarget = False
				filterPerTarget = []
				for f in FiltersByMappingDirection:
					if f["direction"] == sourceSys + ":" + target:
						filterPerTarget = copy.deepcopy(f["filters"])

				for filterV in filterPerTarget:
					if filterV["Sheet"] == sheetID:
						cellVal = str(
							sheet.cell(rownum, getColNumberByName(filterV["Column"], filterV["Sheet"])).value).strip()
						if cellVal == filterV["value"]:
							passByTarget = True
				# get columns by mapping direction
				columns = []
				for colsItem in columnsObj:
					if colsItem["direction"] == sourceSys + ":" + target:
						columns = copy.deepcopy(colsItem["columns"])
				# get columns by sheet
				colsBySheet = {}
				for c in columns:
					if c["Sheet"] == sheetID:
						colsBySheet = copy.deepcopy(c)

				if passByTarget and passBySheet:
					child = str(sheet.cell(rownum, getColNumberByName(colsBySheet["child"], sheetName)).value).replace(
						"\n", " ").strip()
					parent = str(
						sheet.cell(rownum, getColNumberByName(colsBySheet["parent"], sheetName)).value).replace("\n",
																												" ").strip()
					if colsBySheet.get("type") != None:
						typeV = str(
							sheet.cell(rownum, getColNumberByName(colsBySheet["type"], sheetName)).value).replace("\n",
																												  " ").strip()

					findMappingItem = next((item for item in mapObj["map"] if item["sourceRelationshipType"] == typeV),
										   False)

					# create reference object if exist
					newItemSourceType = ""
					thisNewItem = ""
					isReferenceObject = False
					refObjName = ""  # referene attribute name
					if "reference" in parent:
						isReferenceObject = True
						thisNewItem = "parent"
						newItemSourceType = str(sheet.cell(rownum, getColNumberByName(colsBySheet["target_parent"],
																					  sheetName)).value).replace("\n",
																												 " ").strip()
						mappingItem["parentItemType"] = newItemSourceType
						mappingItem["childItemType"] = child
						refObjName = parent.split(":")[1].replace(">", " ").strip()
					if "reference" in child:
						isReferenceObject = True
						thisNewItem = "child"
						newItemSourceType = str(sheet.cell(rownum, getColNumberByName(colsBySheet["target_child"],
																					  sheetName)).value).replace("\n",
																												 " ").strip()
						mappingItem["childItemType"] = newItemSourceType
						mappingItem["parentItemType"] = parent
						refObjName = child.split(":")[1].replace(">", " ").strip()

					mappingItemRefObj = {}

					if isReferenceObject == True:
						# get reference object details
						refObjRequired = str(
							sheet.cell(rownum, getColNumberByName(colsBySheet["required"], sheetName)).value).replace(
							"\n", " ").strip()  # required for all CI types
						refObjItemAttributes = str(sheet.cell(rownum,
															  getColNumberByName(colsBySheet["ref_obj_attributes"],
																				 sheetName)).value).replace("\n",
																											" ").strip().split(
							";")
						mappingItemRefObj["datatype"] = "object"
						mappingItemRefObj["elementType"] = "refObject"
						mappingItemRefObj["name"] = refObjName
						mappingItemRefObj["targets"] = []
						refObj = {}
						refObj["solution"] = target
						# endpoint Servicenow to OMI
						if target == "omi-mvp-job1" or target == "ucmdb-core-ams":
							refObj["endpoint"] = "properties"
						refObj["mapType"] = "refItemAndRelation"
						refObj["required"] = True if refObjRequired == "Y" else False
						refObj["item"] = {}
						refObj["item"]["attributes"] = []
						for itemA in refObjItemAttributes:
							itemAttributes = itemA.split(",")
							refItem = {}
							for a in itemAttributes:
								if "source_attribute:" in a:
									refItem["sourceName"] = a.replace("source_attribute:", " ").strip()
								if "datatype:" in a:
									refItem["datatype"] = a.replace("datatype:", " ").strip()
								if "target_attribute:" in a:
									refItem["targetName"] = a.replace("target_attribute:", " ").strip()
								if "target_value:" in a:
									refItem["targetValue"] = a.replace("target_value:", " ").strip()
								if "existingInRef:" in a:
									exisinRef = a.replace("existingInRef:", " ").strip()
									refItem["existingInRef"] = True if exisinRef == "Y" else False
								if "existingInPri:" in a:
									exisinRef = a.replace("existingInPri:", " ").strip()
									refItem["existingInPri"] = True if exisinRef == "Y" else False
								if "primaryName:" in a:
									refItem["primaryName"] = a.replace("primaryName:", " ").strip()
							refObj["item"]["attributes"].append(refItem)
						refObj["relationship"] = {}
						refObj["relationship"]["sourceRelationshipIdentifier"] = refObjName + "Reference"
						refObj["relationship"]["newItemSourceType"] = newItemSourceType
						refObj["relationship"]["thisNewItem"] = thisNewItem

						mappingItemRefObj["targets"].append(refObj)

						mappingItem["sourceRelationshipType"] = refObjName + "Reference"
					else:
						mappingItem["parentItemType"] = parent
						mappingItem["childItemType"] = child

					if colsBySheet.get("type") != None and isReferenceObject == False:
						mappingItem["sourceRelationshipType"] = typeV
					else:
						mappingItem["sourceRelationshipType"] = refObjName + "Reference"
					mappingItem["targets"] = []
					# Exception ESL relationships

					targetItem["solution"] = target
					if target == "esl":
						# print('subObject')
						mappingItem["elementType"] = "subObject"
						targetItem["endpoint"] = {}
						targetItem["endpoint"]["key"] = "sub-objects"
						targetItem["endpoint"]["keyType"] = "array"
						targetItem["endpoint"]["removeTransformedChild"] = True

						attributesToIDrel = str(sheet.cell(rownum, getColNumberByName(
							colsBySheet["attributes_rel_target"].strip(), sheetName)).value).replace("\n", " ").split(
							";")
						targetItem["attributes"] = []
						# print(attributesToIDrel)
						for attr in attributesToIDrel:
							# attr_name= , datatype= , mapping details separated by *
							attrMapDetails = attr.split(",")
							name = ""
							datatype = ""
							mapDetails = ""
							required = ""
							target_attr = ""
							for det in attrMapDetails:
								if "attr_name=" in det:
									# print(det)
									name = det.replace("attr_name=", " ").strip()
								elif "datatype=" in det:
									datatype = det.replace("datatype=", " ").strip()
								elif "required=" in det:
									required = det.replace("required=", " ").strip()
								elif "target_name=" in det:
									target_attr = det.replace("target_name=", " ").strip()
								else:
									mapDetails = det.strip()
							attrItem = getMaptype(mapDetails, name, False, datatype, sourceSys + ":" + target,
												  columnsObj, required, target_attr)
							if bool(attrItem):
								targetItem["attributes"].append(attrItem)

					else:
						if sourceSys == "esl":
							mappingItem["elementType"] = "eslRelationship"
						else:
							mappingItem["elementType"] = "relationship"
						targetItem["endpoint"] = {}
						targetItem["endpoint"]["include"] = {}

						if colsBySheet.get("childIdKey_target") != None:
							childIdKey = colsBySheet["childIdKey_target"].strip()
							targetItem["childIdKey"] = childIdKey

						if colsBySheet.get("include_child") != None:
							include_child = colsBySheet["include_child"].strip()
							targetItem["endpoint"]["include"][
								"child"] = True if include_child.strip() == "True" else False

						if colsBySheet.get("include_parent") != None:
							include_parent = colsBySheet["include_parent"].strip()
							targetItem["endpoint"]["include"][
								"parent"] = False if include_parent.strip() == "False" else True

						if colsBySheet.get("target_key") != None:
							target_key = colsBySheet["target_key"].strip()
							targetItem["endpoint"]["key"] = target_key

						if colsBySheet.get("target_keytype") != None:
							target_keytype = colsBySheet["target_keytype"].strip()
							targetItem["endpoint"]["keyType"] = target_keytype

						if colsBySheet.get("target_location") != None:
							target_location = colsBySheet["target_location"].strip()
							targetItem["endpoint"]["location"] = target_location

						if colsBySheet.get("type_target") != None:
							type_target = colsBySheet["type_target"].strip()
							targetItem["endpoint"]["type"] = type_target

						if colsBySheet.get("parentIdKey_target") != None:
							parentIdKey_target = colsBySheet["parentIdKey_target"].strip()
							targetItem["parentIdKey"] = parentIdKey_target

						if colsBySheet.get("relationshipType_target") != None:
							relationshipType_target = str(sheet.cell(rownum, getColNumberByName(
								colsBySheet["relationshipType_target"], sheetName)).value).replace("\n", " ").strip()
							targetItem["relationshipType"] = relationshipType_target

						if colsBySheet.get("typeKey_target") != None:
							typeKey_target = colsBySheet["typeKey_target"].strip()
							targetItem["typeKey"] = typeKey_target

						if colsBySheet.get("second_rel") != None:
							# print(colsBySheet)
							secondRel = str(sheet.cell(rownum, getColNumberByName(colsBySheet["second_rel"],
																				  sheetName)).value).replace("\n",
																											 " ").strip()
							if secondRel != 'N' and secondRel != '':
								secondRel = secondRel.split(";")
								targetItem["secondRelationship"] = {}
								targetItem["secondRelationship"]["relationshipType"] = secondRel[0].strip()
								targetItem["secondRelationship"]["order"] = secondRel[1].strip()

					if bool(mappingItemRefObj):
						if findMappingSourceSysRoot == False:
							mapObj["map"].append(mappingItemRefObj)
						else:
							findMappingSourceSysRoot["mapping"]["map"].append(mappingItemRefObj)

					if findMappingItem != False:
						if bool(targetItem):
							findTarget = next((item for item in findMappingItem["targets"] if
											   item["relationshipType"] == targetItem["relationshipType"]), False)
							if findTarget == False:
								findMappingItem["targets"].append(targetItem)
					else:
						if bool(targetItem):
							mappingItem["targets"].append(targetItem)

				if bool(mappingItem) and bool(mappingItem["targets"]):
					if findMappingSourceSysRoot == False:
						mapObj["map"].append(mappingItem)
					else:
						# Exception ESL relationships Ci types with concatenated values
						if ";" in child:
							# Create the same relationship for each posible CI type
							childValues = child.split(";")
							for c in childValues:
								if c.strip() != "":
									itemForEachChild = copy.deepcopy(mappingItem)
									itemForEachChild["childItemType"] = c.strip()
									findMappingSourceSysRoot["mapping"]["map"].append(itemForEachChild)

						else:
							findExistingRelInMapping = False
							for item in findMappingSourceSysRoot["mapping"]["map"]:
								if item["elementType"] == "relationship":
									if item["childItemType"] == mappingItem["childItemType"] and item[
										"sourceRelationshipType"] == mappingItem["sourceRelationshipType"] and item[
										"parentItemType"] == mappingItem["parentItemType"]:
										findExistingRelInMapping = True

							# add the same relationship for each child ci
							if findExistingRelInMapping == False:
								childOfParent = []
								childOfChild = []
								hasChildCI = False
								if parent in hasChilds:
									parenthood = next((item["root"] for item in ciParentList if item["ci"] == parent),
													  False)
									childOfParent = getCITypesByParent(parenthood, columnsObj, sourceSys + ":" + target)
									hasChildCI = True
								if child in hasChilds:
									parenthood = next((item["root"] for item in ciParentList if item["ci"] == child),
													  False)
									childOfChild = getCITypesByParent(parenthood, columnsObj, sourceSys + ":" + target)
									hasChildCI = True
								if hasChildCI:
									if len(childOfParent) > 1 and len(childOfChild) > 1:
										for ciP in childOfParent:
											for ciC in childOfChild:
												# if "virtual_center" in ciC and mappingItem["elementType"] == "eslRelationship":
												# print(mappingItem)
												childObj = copy.deepcopy(mappingItem)
												childObj["parentItemType"] = ciP
												childObj["childItemType"] = ciC
												findExistingRelInMapping = False
												for item in findMappingSourceSysRoot["mapping"]["map"]:
													if item["elementType"] == "relationship":
														if item["childItemType"] == childObj["childItemType"] and item[
															"sourceRelationshipType"] == childObj[
															"sourceRelationshipType"] and item["parentItemType"] == \
																childObj["parentItemType"]:
															findExistingRelInMapping = True
												if findExistingRelInMapping == False:
													findMappingSourceSysRoot["mapping"]["map"].append(childObj)
									else:
										if len(childOfParent) > 1:
											for ci in childOfParent:
												childObj = copy.deepcopy(mappingItem)
												childObj["parentItemType"] = ci
												findExistingRelInMapping = False
												# print("child "+mappingItem["parentItemType"])
												for item in findMappingSourceSysRoot["mapping"]["map"]:
													if item["elementType"] == "relationship":
														if item["childItemType"] == childObj["childItemType"] and item[
															"sourceRelationshipType"] == childObj[
															"sourceRelationshipType"] and item["parentItemType"] == \
																childObj["parentItemType"]:
															findExistingRelInMapping = True
												if findExistingRelInMapping == False:
													findMappingSourceSysRoot["mapping"]["map"].append(childObj)
										elif len(childOfChild) > 1:
											for ci in childOfParent:
												childObj = copy.deepcopy(mappingItem)
												childObj["childItemType"] = ci
												findExistingRelInMapping = False
												for item in findMappingSourceSysRoot["mapping"]["map"]:
													if item["elementType"] == "relationship":
														if item["childItemType"] == childObj["childItemType"] and item[
															"sourceRelationshipType"] == childObj[
															"sourceRelationshipType"] and item["parentItemType"] == \
																childObj["parentItemType"]:
															findExistingRelInMapping = True
												if findExistingRelInMapping == False:
													findMappingSourceSysRoot["mapping"]["map"].append(childObj)
								else:
									findMappingSourceSysRoot["mapping"]["map"].append(mappingItem)


def processSheet(sheetName, sheetID, sourceSys, integration, mappingRootObj, columnsObj, targets):
	##print("processing attr method")
	sheet = workbook.sheet_by_name(sheetName.strip())
	mapBySource = {}
	mapObj = {}
	mapObj["map"] = []
	findMappingSourceSysRoot = next((item for item in mappingRootObj if item["source"] == sourceSys + integration),
									False)
	# print(json.dumps(findMappingSourceSysRoot,indent=4, sort_keys=False))

	for rownum in range(sheet.nrows):
		# genaral filter by sheet
		for target in targets:
			mappingItem = {}
			targetItem = {}
			# filter by target
			passBySheet = False
			for filterBySheet in FiltersBySheet[sheetID]:
				cellVal = str(sheet.cell(rownum, getColNumberByName(filterBySheet["Column"], sheetName)).value).strip()
				if cellVal == filterBySheet["value"]:
					passBySheet = True
				else:
					passBySheet = False

			# filter by direction
			passByTarget = False
			filterPerTarget = []
			for f in FiltersByMappingDirection:
				if f["direction"] == sourceSys + ":" + target:
					filterPerTarget = copy.deepcopy(f["filters"])

			for filterV in filterPerTarget:
				if filterV["Sheet"] == sheetID:
					cellVal = str(
						sheet.cell(rownum, getColNumberByName(filterV["Column"], filterV["Sheet"])).value).strip()
					if cellVal == filterV["value"]:
						passByTarget = True
			# get columns by mapping direction
			columns = []
			for colsItem in columnsObj:
				if colsItem["direction"] == sourceSys + ":" + target:
					columns = copy.deepcopy(colsItem["columns"])
			# get columns by sheet
			colsBySheet = {}
			for c in columns:
				if c["Sheet"] == sheetID:
					colsBySheet = copy.deepcopy(c)

			itemName = str(sheet.cell(rownum, getColNumberByName(colsBySheet["name"], sheetName)).value).replace("\n",
																												 " ").strip()
			# exception display_label
			if 'discovered_product_name' in itemName:
				itemName = 'discovered_product_name'

			if passByTarget and passBySheet:
				# create source part
				mappingItem["name"] = itemName
				mappingItem["elementType"] = "attribute"
				mappingItem["datatype"] = str(
					sheet.cell(rownum, getColNumberByName(colsBySheet["datatype"], sheetName)).value).strip()
				mappingItem["targets"] = []

				# get target details
				mapType = str(sheet.cell(rownum, getColNumberByName(colsBySheet["mapType"], sheetName)).value).strip()
				endpoint = str(sheet.cell(rownum, getColNumberByName(colsBySheet["endpoint"], sheetName)).value).strip()
				name = str(sheet.cell(rownum, getColNumberByName(colsBySheet["target_name"], sheetName)).value).strip()
				datatype = str(sheet.cell(rownum, getColNumberByName(colsBySheet["target_datatype"],
																	 sheetName)).value).strip().lower()
				required = str(sheet.cell(rownum, getColNumberByName(colsBySheet["required"], sheetName)).value).strip()
				calculations = str(
					sheet.cell(rownum, getColNumberByName(colsBySheet["calculations"], sheetName)).value).strip()
				dataformat = 'N/A'
				if colsBySheet.get("target_format") != None:
					dataformat = str(
						sheet.cell(rownum, getColNumberByName(colsBySheet["target_format"], sheetName)).value).strip()
				mapTypeDetails = {}
				if calculations != "N":
					calc_details = calculations.split(";")
					for idx, c in enumerate(calc_details):
						if findMappingSourceSysRoot != False:
							findMappingItem = next((item for item in findMappingSourceSysRoot["mapping"]["map"] if
													item["name"] == itemName), False)
						else:
							findMappingItem = next((item for item in mapObj["map"] if item["name"] == itemName), False)

						details = c.strip().split("=")
						# print(itemName)
						# print(details)
						if len(details) > 1:
							detailsKey = details[0].strip()
							direction = sourceSys + ":" + target

							# get maptype details
							# print(detailsKey)
							# print(details)
							mapTypeDetails = getMaptypeDetails(mappingItem, itemName, detailsKey, details, direction,
															   columnsObj, name, endpoint, datatype, required,
															   dataformat)

						requiredBoolVal = True if required == "Y" or required == "Y - All CI Types" else False
						# ESL exception for Solution name only going to Installed software CI
						targetAttributeNames = str(sheet.cell(rownum, getColNumberByName(colsBySheet["target_name"],
																						 sheetName)).value).strip().split(
							";")
						if len(targetAttributeNames) == len(calc_details):
							name = targetAttributeNames[idx].strip()
						mapTypes = mapType.split(";")
						if len(mapTypes) == len(calc_details):
							mapType = mapTypes[idx].strip()
						# exception display_label
						if 'discovered_product_name' in name and 'servicenow:' in direction:
							name = 'discovered_product_name'
						targetItem = getTargetItemByMapType(mapType, name, requiredBoolVal, target, endpoint, datatype,
															mapTypeDetails, dataformat)
						if colsBySheet.get("priority_source") != None:
							if str(sheet.cell(rownum, getColNumberByName(colsBySheet["priority_source"],
																		 sheetName)).value).strip() == 'Y':
								targetItem['useCorrelationUniqueId'] = True
						cloneItem = {}

						if findMappingItem != False:
							if bool(targetItem):

								if len(findMappingItem["targets"]) == 2:
									duplicate = False
									for item in findMappingItem["targets"]:
										if sorted(item.items()) == sorted(targetItem.items()):
											duplicate = True
									if duplicate == False:
										cloneItem = copy.deepcopy(findMappingItem)
										cloneItem["targets"] = []
										cloneItem["targets"].append(targetItem)
										findMappingSourceSysRoot["mapping"]["map"].append(cloneItem)
								else:
									targetSolutionExists = False
									for t in findMappingItem["targets"]:
										if t["solution"] == targetItem["solution"]:
											targetSolutionExists = True
									if targetSolutionExists == True:
										# clone and add
										cloneItem = copy.deepcopy(findMappingItem)
										cloneItem["targets"] = []
										cloneItem["targets"].append(targetItem)
										# add only if it does not exist already
										itemExists = False
										ExistsWithNameAndTarget = False
										for item in findMappingSourceSysRoot["mapping"]["map"]:
											if item["name"] == cloneItem["name"]:
												for t in item["targets"]:
													if t["name"] == targetItem["name"]:
														ExistsWithNameAndTarget = True

										if ExistsWithNameAndTarget == False:
											findMappingSourceSysRoot["mapping"]["map"].append(cloneItem)
									else:
										findMappingItem["targets"].append(targetItem)
						else:
							if bool(targetItem):
								mappingItem["targets"].append(targetItem)

						if bool(mappingItem) and bool(mappingItem["targets"]) and findMappingItem == False:
							if findMappingSourceSysRoot == False:
								mapObj["map"].append(mappingItem)
							else:
								findMappingSourceSysRoot["mapping"]["map"].append(mappingItem)


def getTargetItemByMapType(mapType, name, requiredBoolVal, target, endpoint, datatype, mapTypeDetails, dataformat):
	targetItem = {}
	if mapType == "direct":
		targetItem = {
			"mapType": mapType,
			"name": name,
			"required": requiredBoolVal,
			"solution": target,
			"endpoint": endpoint,
			"datatype": datatype
		}
		if dataformat != 'N/A':
			targetItem['dataFormat'] = dataformat
	elif mapType == "list":
		targetItem = {
			"mapType": mapType,
			"name": name,
			"required": requiredBoolVal,
			"solution": target,
			"endpoint": endpoint,
			"datatype": datatype,
			"options": "List mapType requires options" if mapTypeDetails.get("options") == None else mapTypeDetails.get(
				"options")
		}
		if dataformat != 'N/A':
			targetItem['dataFormat'] = dataformat
	elif mapType == "derived":
		targetItem = {
			"mapType": mapType,
			"name": name,
			"required": requiredBoolVal,
			"solution": target,
			"endpoint": endpoint,
			"datatype": datatype,
			"relation": "Direct-derived mapType requires relation" if mapTypeDetails.get(
				"relation") == None else mapTypeDetails.get("relation")
		}
		if dataformat != 'N/A':
			targetItem['dataFormat'] = dataformat
	elif mapType == "concatenation":
		targetItem = {
			"mapType": mapType,
			"name": name,
			"required": requiredBoolVal,
			"solution": target,
			"endpoint": endpoint,
			"datatype": datatype,
			"values": "Concatenation mapType requires values" if mapTypeDetails.get(
				"values") == None else mapTypeDetails.get("values")
		}
		if dataformat != 'N/A':
			targetItem['dataFormat'] = dataformat
	elif mapType == "calculation":
		targetItem = {
			"mapType": mapType,
			"name": name,
			"required": requiredBoolVal,
			"solution": target,
			"endpoint": endpoint,
			"datatype": datatype,
			"formula": "Calculation mapType requires formula" if mapTypeDetails.get(
				"formula") == None else mapTypeDetails.get("formula"),
			"context": "Calculation mapType requires context" if mapTypeDetails.get(
				"context") == None else mapTypeDetails.get("context"),
			"matchChildItemType": "Calculation mapType requires matchChildItemType" if mapTypeDetails.get(
				"matchChildItemType") == None else mapTypeDetails.get("matchChildItemType"),
			"matchRelType": "Calculation mapType requires matchRelType" if mapTypeDetails.get(
				"matchRelType") == None else mapTypeDetails.get("matchRelType")
		}
		if dataformat != 'N/A':
			targetItem['dataFormat'] = dataformat
	elif mapType == "conditional":
		targetItem = {
			"mapType": mapType,
			"name": name,
			"required": requiredBoolVal,
			"solution": target,
			"endpoint": endpoint,
			"datatype": datatype,
			"condition": "Conditional mapType requires condition" if mapTypeDetails.get(
				"condition") == None else mapTypeDetails.get("condition")
		}
		if dataformat != 'N/A':
			targetItem['dataFormat'] = dataformat
	else:
		targetItem = {
			"mapType": "mapType",
			"name": name,
			"required": requiredBoolVal,
			"solution": target,
			"endpoint": endpoint,
			"datatype": datatype
		}
		if dataformat != 'N/A':
			targetItem['dataFormat'] = dataformat

	return targetItem


def getMaptypeDetails(mappingItem, itemName, detailsKey, details, direction, columnsObj, name, endpoint, datatype,
					  required, dataformat):
	mapTypeDetails = {}
	if detailsKey == "options":
		options = []
		optionsDetails = details[1].split(",")
		# Exception ESL non sheet values
		if "{" in details[1]:
			mappingItem["name"] = optionsDetails[0].strip()
			preLoadedValues = optionsDetails[1].replace("{", " ").replace("}", " ").strip().split("*")
			options = getListOptions(None, None, None, None, None, None, False, [], preLoadedValues)
		else:
			# print(optionsDetails)
			excludeL = []
			if len(optionsDetails) > 6:
				excludeL = optionsDetails[6].replace("exclude:", " ").strip().split("|")
			# print(excludeL)
			options = getListOptions(optionsDetails[0], optionsDetails[1], optionsDetails[2], optionsDetails[3],
									 optionsDetails[4], direction, False, excludeL, [])
		mapTypeDetails[detailsKey] = options
	elif detailsKey == "relation":
		relationDetails = details[1].split(",")
		relation = {}
		relation["idAttribute"] = relationDetails[0].strip()
		relation["options"] = []
		relation["options"] = getOptionsDirectDerived(relationDetails[1], relationDetails[4], direction, columnsObj)
		relation["type"] = relationDetails[2].strip()
		relation["current"] = "parent" if relationDetails[3] == "child" else "child"
		# print("derived")
		relation["map"] = getMaptype(relationDetails[5], relationDetails[4], endpoint, datatype, direction, columnsObj,
									 required, dataformat)
		mapTypeDetails[detailsKey] = relation
	elif detailsKey == "values":
		valuesA = details[1].split(",")
		values = []
		for val in valuesA:
			val = val.replace('{eq}', '=').replace('{semicolon}', ';')
			if "constant:" in val:
				val = val.replace("constant:", " ").strip()
				if "true" in val:
					val = True
				elif "space" in val:
					val = " "
				values.append({"type": "constant", "value": val})
			else:
				values.append({"type": "variable", "value": val.strip()})
		mapTypeDetails[detailsKey] = values
	elif detailsKey == "formula":
		formulaDetails = details[1].split(",")
		mapTypeDetails["formula"] = formulaDetails[0].strip()
		mapTypeDetails["context"] = formulaDetails[1].strip()
		mapTypeDetails["matchRelType"] = formulaDetails[2].strip()
		mapTypeDetails["matchChildItemType"] = formulaDetails[3].strip()
	elif detailsKey == "condition":
		conditionDetails = details[1].split(",")

		conditionObject = {}
		for c in conditionDetails:
			if "if:" in c:
				conditionObject["if"] = c.replace("if:", " ").strip()
			if "in:" in c:
				exception = False
				if "esl:" in direction and ("Instance =>" in itemName or "N/A" in itemName):
					exception = c.replace("in:", " ").strip()
					conditionObject["in"] = getCItypesList(c.replace("in:", " ").strip(), direction, columnsObj,
														   exception)
				elif "*" in c:
					conditionObject["in"] = c.replace("in:", " ").strip().split("*")
					for idx, el in enumerate(conditionObject["in"]):
						if el == '':
							del conditionObject["in"][idx]
						else:
							conditionObject["in"][idx] = el.strip()
				else:
					CIsList = getCItypesList(c.replace("in:", " ").strip(), direction, columnsObj, exception)
					if len(CIsList) == 0:
						conditionObject["in"] = [c.replace("in:", " ").strip()]
					else:
						conditionObject["in"] = CIsList
			if "then:" in c:
				conditionObject["then"] = {}
				conditionObject["then"] = getMaptype(c.replace("then:", " ").strip(), name, endpoint, datatype,
													 direction, columnsObj, False, dataformat)
			if "else:" in c:
				conditionObject["else"] = {}
				conditionObject["else"] = getMaptype(c.replace("else:", " ").strip(), name, endpoint, datatype,
													 direction, columnsObj, False, dataformat)
		mapTypeDetails["condition"] = conditionObject
	else:
		mapTypeDetails[detailsKey] = details[1]

	return mapTypeDetails


def getMaptype(mapTypeDetails, name, endpoint, datatype, direction, columnsObj, required, dataformat, *target_attr):
	##print("getMaptype method")
	mapTypeObject = {}
	if "values:" in mapTypeDetails:
		values = mapTypeDetails.replace("values:", " ").strip().split("*")
		mapTypeObject["mapType"] = "concatenation"
		mapTypeObject["name"] = name
		if endpoint != False:
			mapTypeObject["endpoint"] = endpoint
		mapTypeObject["datatype"] = datatype
		if dataformat != 'N/A':
			mapTypeObject['dataFormat'] = dataformat
		valuesObj = []
		for v in values:
			if "constant:" in v:
				valuesObj.append({"type": "constant", "value": v.replace("constant:", " ").strip()})
			else:
				valuesObj.append({"type": "variable", "value": v.strip()})
		mapTypeObject["values"] = valuesObj

	elif "relation:" in mapTypeDetails:
		relationObj = {}
		relationDetails = mapTypeDetails.replace("relation:", " ").strip().split("*")
		relationObj["idAttribute"] = relationDetails[0].strip()
		relationObj["options"] = []
		relationObj["options"] = getOptionsDirectDerived(relationDetails[1], relationDetails[4], direction, columnsObj)
		relationObj["type"] = relationDetails[2].strip()
		relationObj["current"] = "parent" if relationDetails[3] == "child" else "child"
		# print(relationDetails)
		# print(relationDetails)
		relationObj["map"] = getMaptype(relationDetails[5].replace("values", "values:").replace("-", "*"), name,
										endpoint, datatype, direction, columnsObj, required, dataformat)
		mapTypeObject["mapType"] = "derived"
		mapTypeObject["relation"] = relationObj

	elif "attribute:" in mapTypeDetails:
		# print(mapTypeDetails)
		# print(name)
		mapTypeObject["mapType"] = "direct"
		targetAttrName = name
		if "attribute:" in mapTypeDetails:
			name = mapTypeDetails.replace("attribute:", " ").strip()
			if name != "":
				mapTypeObject["name"] = name
		else:
			mapTypeObject["name"] = name
		if ":esl" in direction:
			mapTypeObject["attribute"] = targetAttrName
			mapTypeObject["datatype"] = datatype
			mapTypeObject["required"] = True if required == "Y" else False
			if dataformat != 'N/A':
				mapTypeObject['dataFormat'] = dataformat
	elif "skip" in mapTypeDetails:
		mapTypeObject["mapType"] = "skip"
	elif "list:" in mapTypeDetails:
		print(mapTypeDetails)
		mapTypeObject["mapType"] = "list"
		listDetails = mapTypeDetails.replace("list:", " ").strip().split("*")
		if target_attr:
			mapTypeObject["attribute"] = name
			mapTypeObject["name"] = target_attr[0]
		attribute = "False"
		if len(listDetails) > 5 and listDetails[5] != '':
			attribute = listDetails[5]
			if attribute != "False" and attribute != '':
				mapTypeObject["idAttribute"] = attribute.strip()
		excludeItems = []
		if len(listDetails) > 6:
			excludeItems = listDetails[6].replace("exclude:", " ").strip().split("|")
		# print(excludeItems)
		mapTypeObject["options"] = getListOptions(listDetails[0], listDetails[1], listDetails[2], listDetails[3],
												  listDetails[4], direction, attribute, excludeItems, [])
	return mapTypeObject


def getCItypesList(ciType, mapDirection, columnsObj, exception):
	##print("getCItypesList method")
	options = []
	sheet = workbook.sheet_by_name(CITypesSheet.strip())
	sheetID = "CI Types"
	for rownum in range(sheet.nrows):
		passBySheet = False
		for filterBySheet in FiltersBySheet["CI Types"]:
			cellVal = str(sheet.cell(rownum, getColNumberByName(filterBySheet["Column"], CITypesSheet)).value).strip()
			if cellVal == filterBySheet["value"]:
				passBySheet = True
		if passBySheet:
			# get columns by mapping direction
			columns = []
			for colsItem in columnsObj:
				if colsItem["direction"] == mapDirection:
					columns = copy.deepcopy(colsItem["columns"])
			# get columns by sheet
			colsBySheet = {}
			for c in columns:
				if c["Sheet"] == sheetID:
					colsBySheet = copy.deepcopy(c)
			if colsBySheet.get("parenthood") != None:
				if exception == False:
					# Exception: ESL CI type based on UCMDB CI types name
					if "esl:" in mapDirection and ciType[0].isupper() == False:
						citypename = str(sheet.cell(rownum, getColNumberByName(colsBySheet["target_name"],
																			   CITypesSheet)).value).strip()
					else:
						citypename = str(
							sheet.cell(rownum, getColNumberByName(colsBySheet["name"], CITypesSheet)).value).strip()
						if "|" in citypename:
							citypename = citypename.split("|")[0].strip()

					if ">" in citypename:
						displayValue = citypename.split(">")
						citypename = displayValue[len(displayValue) - 1]

					if citypename == ciType:
						parenthood = str(sheet.cell(rownum, getColNumberByName(colsBySheet["parenthood"],
																			   CITypesSheet)).value).strip()
						ciTypesByParent = getCITypesByParent(parenthood, columnsObj, mapDirection)
						if len(ciTypesByParent) > 1 and parenthood != "":
							for ci in ciTypesByParent:
								if "\n" in ci:
									cis = ci.split("\n")
									for c in cis:
										findCIinOptions = next((item for item in options if item == c), False)
										if findCIinOptions == False:
											options.append(c)
								else:
									findCIinOptions = next((item for item in options if item == ci), False)
									if findCIinOptions == False:
										options.append(ci)
						else:
							if citypename != "":
								findCIinOptions = next((item for item in options if item == citypename), False)
								if findCIinOptions == False:
									options.append(citypename)
				else:
					# filter CI Types by exception
					citypename = str(
						sheet.cell(rownum, getColNumberByName(colsBySheet["name"], CITypesSheet)).value).strip()
					if exception in citypename and citypename != "":
						options.append(citypename)

	return options


def getCITypesByParent(parenthood, columnsObj, mapDirection):
	##print("getCITypesByParent rel method")
	cis = []
	sheet = workbook.sheet_by_name(CITypesSheet.strip())
	sheetID = "CI Types"
	for rownum in range(sheet.nrows):
		# get columns by mapping direction
		columns = []
		for colsItem in columnsObj:
			if colsItem["direction"] == mapDirection:
				columns = copy.deepcopy(colsItem["columns"])
		# get columns by sheet
		colsBySheet = {}
		for c in columns:
			if c["Sheet"] == sheetID:
				colsBySheet = copy.deepcopy(c)
		if colsBySheet.get("parenthood") != None:
			parenthoodVal = str(
				sheet.cell(rownum, getColNumberByName(colsBySheet["parenthood"], CITypesSheet)).value).strip()
			citypename = str(sheet.cell(rownum, getColNumberByName(colsBySheet["name"], CITypesSheet)).value).strip()
			perenthoodchild = parenthood + ">"
			if parenthood == parenthoodVal or perenthoodchild in parenthoodVal:
				cis.append(citypename)
	return cis


def getParenthoodByCI(ciname, mapDirection, columnsObj):
	##print("getParenthoodByCI rel method")

	sheet = workbook.sheet_by_name(CITypesSheet.strip())
	sheetID = "CI Types"
	parenthoodVal = ""
	for rownum in range(sheet.nrows):
		# get columns by mapping direction
		columns = []
		for colsItem in columnsObj:
			if colsItem["direction"] == mapDirection:
				columns = copy.deepcopy(colsItem["columns"])
		# get columns by sheet
		colsBySheet = {}
		for c in columns:
			if c["Sheet"] == sheetID:
				colsBySheet = copy.deepcopy(c)
		if colsBySheet.get("parenthood") != None:
			citypename = str(sheet.cell(rownum, getColNumberByName(colsBySheet["name"], CITypesSheet)).value).strip()
			if citypename == ciname:
				parenthoodVal = str(
					sheet.cell(rownum, getColNumberByName(colsBySheet["parenthood"], CITypesSheet)).value).strip()
	return parenthoodVal


def getOptionsDirectDerived(ciType, attribute, mapDirection, columnsObj):
	##print("getOptionsDirectDerived rel method")

	options = []
	parent = getParenthoodByCI(ciType.strip(), mapDirection, columnsObj)
	cis = getCITypesByParent(parent, columnsObj, mapDirection)
	for ci in cis:
		options.append({"condition": ci, "attribute": attribute.strip()})

	return options


def getListOptions(source, sourceColumn, sourceValueColumn, targetValueColumn, sheetName, mapDirection, attribute,
				   excludeTarget, preLoadedValues):
	# print("getListOptions rel method")
	options = []
	if len(preLoadedValues) > 0:
		# Exception ESL list items not in sheets
		for item in preLoadedValues:
			values = item.split(">")
			options.append({"sourceValue": values[0].strip(), "targetValue": values[1].strip()})
	else:
		sheet = workbook.sheet_by_name(sheetName.strip())
		for rownum in range(sheet.nrows):
			passBySheet = False
			for filterBySheet in FiltersBySheet[sheetName.strip()]:
				cellVal = str(
					sheet.cell(rownum, getColNumberByName(filterBySheet["Column"], sheetName.strip())).value).strip()
				if cellVal == filterBySheet["value"]:
					passBySheet = True
			if passBySheet:
				# print(source+"-"+ sourceColumn+"-"+ sourceValueColumn+"-"+ targetValueColumn+"-"+ sheetName+"-"+ mapDirection+"-"+str(attribute)+"-"+str(excludeTarget))
				sourceV = str(
					sheet.cell(rownum, getColNumberByName(sourceColumn.strip(), sheetName.strip())).value).strip()
				# print(source.strip() + ' ==? ' +sourceV)
				# print(source.strip() in sourceV)
				if source.strip() in sourceV:
					sourceValue = str(sheet.cell(rownum, getColNumberByName(sourceValueColumn.strip(),
																			sheetName.strip())).value).strip()
					if ">" in sourceValue:
						displayValue = sourceValue.split(">")
						sourceValue = displayValue[len(displayValue) - 1]
					targetValue = str(sheet.cell(rownum, getColNumberByName(targetValueColumn.strip(),
																			sheetName.strip())).value).strip()
					# print(sourceValue)
					# Exception SN unix mapping
					if "Use=" in sourceValue:
						sourceValues = sourceValue.split(";")
						targetValues = targetValue.split("\n")
						for idx, val in enumerate(sourceValues):
							if "Use=" not in val:
								valList = val.split("=")[1].split(",")
								for v in valList:
									options.append(
										{"sourceValue": v.strip(), "targetValue": targetValues[idx - 1].strip()})
					# Exception node_role check for net_devices ITAM
					elif " contains " in sourceValue and 'N/A' not in targetValue.strip():
						# print(targetValue)
						# print(sourceValue)
						# print(targetValue)
						sourceVal = sourceValue.split('contains')[1].strip()
						options.append({"sourceValue": sourceVal.strip(), "targetValue": targetValue.strip()})
					elif "\n" in sourceValue and 'N/A' not in targetValue.strip():
						sourceValues = sourceValue.split("\n")
						for idx, val in enumerate(sourceValues):
							options.append({"sourceValue": val.strip(), "targetValue": targetValue.strip()})
					elif "\n" not in targetValue and 'N/A' not in targetValue.strip():
						# Exception ESL receives Object Type not concatenated
						if "|" in targetValue and ":esl" in mapDirection:
							getTargetValue = targetValue.split("|")
							targetValue = getTargetValue[0]
						# Exception ESL System Type Calculation
						if "System Type =" in targetValue:
							# print(targetValue)
							targetValue = targetValue.replace("System Type =", " ").strip()
						if attribute != "False" and attribute != False and len(excludeTarget) == 0:
							if 'true' in sourceValue:
								sourceValue = True
							if 'false' in sourceValue:
								sourceValue = False
							if 'true' in targetValue:
								targetValue = True
							if 'false' in targetValue:
								targetValue = False
							options.append({"sourceValue": source.strip(), "targetValue": targetValue.strip()})
						else:
							# Exception ESL outbound will send concatenated values for root_class
							# print("---------------------------------------------")
							# print(excludeTarget)
							# print("---------------------------------------------")
							# print(targetValue)
							# print(sourceValue)
							# print(sourceValue.strip() not in excludeTarget and targetValue.strip() not in excludeTarget and len(excludeTarget)>0 and sourceValue.strip() != "")
							if "esl:" in mapDirection:
								if "\n" in sourceValue:
									sourceValues = sourceValue.split("\n")
									for s in sourceValues:
										if 'true' in sourceValue:
											sourceValue = True
										if 'false' in sourceValue:
											sourceValue = False
										if 'true' in targetValue:
											targetValue = True
										if 'false' in targetValue:
											targetValue = False
										options.append({"sourceValue": s.strip(), "targetValue": targetValue.strip()})
								else:
									options.append(
										{"sourceValue": sourceValue.strip(), "targetValue": targetValue.strip()})
							elif (
									sourceValue.strip() not in excludeTarget and targetValue.strip() not in excludeTarget) and len(
									excludeTarget) > 0 and sourceValue.strip() != "" and 'N/A' not in targetValue.strip():
								# print("exclude")
								# print(excludeTarget)
								if 'true' in sourceValue:
									sourceValue = True
								if 'false' in sourceValue:
									sourceValue = False
								if 'true' in targetValue:
									targetValue = True
								if 'false' in targetValue:
									targetValue = False
								options.append({"sourceValue": sourceValue.strip(), "targetValue": targetValue.strip()})
							elif len(excludeTarget) == 0 and 'N/A' not in targetValue.strip():
								if "\n" in sourceValue:
									sources = sourceValue.split("\n")
									for s in sources:
										if 'true' in sourceValue:
											sourceValue = True
										elif 'false' in sourceValue:
											sourceValue = False
										if 'true' in targetValue:
											targetValue = True
										elif 'false' in targetValue:
											targetValue = False
										options.append({"sourceValue": s.strip(), "targetValue": targetValue.strip()})
								else:
									if sourceValue != "":
										sourceValue = sourceValue.strip()
										targetValue = targetValue.strip()
										if 'true' in sourceValue:
											sourceValue = True
										elif 'false' in sourceValue and sourceValue != True:
											sourceValue = False
										if 'true' in targetValue:
											targetValue = True
										elif 'false' in targetValue and targetValue != True:
											targetValue = False
										options.append({"sourceValue": sourceValue, "targetValue": targetValue})

					# Exception ESL receives Object Type not concatenated
					elif "|" in targetValue and ":esl" in mapDirection and 'N/A' not in targetValue.strip():
						getTargetValue = targetValue.split("|")
						targetValue = getTargetValue[0]
						if 'true' in sourceValue:
							sourceValue = True
						if 'false' in sourceValue:
							sourceValue = False
						if 'true' in targetValue:
							targetValue = True
						if 'false' in targetValue:
							targetValue = False
						options.append({"sourceValue": sourceValue.strip(), "targetValue": targetValue.strip()})

	return options


# Open Sheets
log.write("[INFO] - Opening sheets\n")
if bool(multipleAttrSheet):
	# process multiple sheets
	for s in multipleAttrSheet:
		openS = workbook.sheet_by_name(multipleAttrSheet[s].strip())
	# call function to process sheet
	log.write("[INFO] - Process sheet: " + multipleAttrSheet[s].strip() + "\n")
# print("[INFO] - Process sheet: "+ multipleAttrSheet[s].strip())
else:
	# attrSheet = workbook.sheet_by_name(AttributesSheet)
	# process sheet for each source system
	for mapDir in mappingDirections:
		# print(json.dumps(mapDir,indent=4, sort_keys=False))
		processSheet(AttributesSheet, "Attributes", mapDir["source"], mapDir["integration"], mapping_json,
					 columnsToGetValuesBySheetNMapDireccions, mapDir["targets"])

# print("processing Attributes Sheet")

if bool(multipleCITSheet):
	# process multiple sheets
	for s in multipleCITSheet:
		openS = workbook.sheet_by_name(multipleCITSheet[s].strip())
	# call function to process sheet
	log.write("[INFO] - Process sheet: " + multipleCITSheet[s].strip() + "\n")
# print("[INFO] - Process sheet: "+ multipleCITSheet[s].strip())
else:
	citSheet = workbook.sheet_by_name(CITypesSheet)
if bool(multipleRelSheet):
	# process multiple sheets
	for indx, s in enumerate(multipleRelSheet):
		# call function to process sheet
		for mapDir in mappingDirections:
			processRelationshipsSheet(multipleRelSheet["RelSheet" + str(indx)], "Relationships", mapDir["source"],
									  mapDir["integration"], mapping_json, columnsToGetValuesBySheetNMapDireccions,
									  mapDir["targets"], False)

	log.write("[INFO] - Process sheet: " + multipleRelSheet[s].strip() + "\n")
# print("[INFO] - Process sheet: "+ multipleRelSheet[s].strip())
else:
	for mapDir in mappingDirections:
		processRelationshipsSheet(RelationshipsSheet, "Relationships", mapDir["source"], mapDir["integration"],
								  mapping_json, columnsToGetValuesBySheetNMapDireccions, mapDir["targets"], False)

if bool(multipleEnumSheet):
	# process multiple sheets
	for s in multipleEnumSheet:
		openS = workbook.sheet_by_name(multipleEnumSheet[s].strip())
	# call function to process sheet
	log.write("[INFO] - Process sheet: " + multipleEnumSheet[s].strip() + "\n")
# print("[INFO] - Process sheet: "+ multipleEnumSheet[s].strip())
else:
	log.write("[INFO] - Process sheet: " + EnumSheet.strip() + "\n")
# enumSheet = workbook.sheet_by_name(EnumSheet)

##print(json.dumps(mapping_json,indent=4, sort_keys=False))
AllMappings = []
result = {}
# print(json.dumps(UsecasesAndConnectionBySource,indent=4, sort_keys=False))
# create files according to each mapping direction
for mapDir in mappingDirections:
	# print(mapDir)
	targets = mapDir["targets"]
	# filename = "mapping_"+mapDir["source"]+"-"+targets+".txt"
	findMappingByDirection = next(
		(item for item in mapping_json if item["source"] == mapDir["source"] + mapDir["integration"]), False)
	# print(mapDir["source"]+mapDir["integration"])
	# print(targets)
	# print("FOUND!!!!!")
	# print(findMappingByDirection["mapping"]["map"][0])

	findUsecasesBySource = next((item for item in UsecasesAndConnectionBySource if
								 item["source"] == mapDir["source"] and item["integration"] == mapDir["integration"]),
								False)
	# print(json.dumps(findUsecasesBySource,indent=4, sort_keys=False))
	# print(findUsecasesBySource["usecases"])
	# for each usecase add the target details
	if findUsecasesBySource != False:
		mappingItem = {}
		mappingItem = copy.deepcopy(findMappingByDirection["mapping"])
		# print("COPIED!!!!")
		# print(mappingItem["map"][0])
		mappingItem["createdBy"] = "Python script"
		mappingItem["active"] = True
		mappingItem["targetDetails"] = {}
		AllTargets = {}
		for usecase in findUsecasesBySource["usecases"]:
			mappingItem["usecase"] = usecase["usecase"]
			mappingItem["version"] = usecase["version"]
			##print(usecase)
			findFinalMappingBySourceaAndUsecase = next((item for item in AllMappings if
														item["source"] == mapDir["source"] and item["usecase"] ==
														mappingItem["usecase"]), False)
			target = usecase["target"]
			targetDetails = {}
			targetDetails[target] = {}

			findConnectionsBySourceIntegration = next(
				(item.get(mapDir["source"] + ":" + usecase["target"] + "-" + findUsecasesBySource["integration"]) for
				 item in connectionsByDirectionAndIntegration if
				 item.get(mapDir["source"] + ":" + usecase["target"] + "-" + findUsecasesBySource["integration"])),
				False)
			targetDetails[target]["connection"] = {}
			targetDetails[target]["connection"] = findConnectionsBySourceIntegration
			# targetDetails[target]["connection"]["auth"] = usecase["connection"]["auth"]
			# targetDetails[target]["connection"]["basePath"] = usecase["connection"]["basePath"]
			# targetDetails[target]["connection"]["host"] =  usecase["connection"]["host"]
			# targetDetails[target]["connection"]["method"] =  usecase["connection"]["method"]
			# targetDetails[target]["connection"]["port"] =  usecase["connection"]["port"]
			# targetDetails[target]["connection"]["pwd"] =  usecase["connection"]["pwd"]
			# if usecase["connection"].get("syncPath") != None:
			# targetDetails[target]["connection"]["syncPath"] =  usecase["connection"]["syncPath"]

			# targetDetails[target]["connection"]["type"] =  usecase["connection"]["type"]
			# targetDetails[target]["connection"]["user"] =  usecase["connection"]["user"]

			targetDetails[target]["filter"] = {}
			targetDetails[target]["include"] = {}
			targetDetails[target]["filter"] = usecase["targetDetails"]["filter"]
			targetDetails[target]["include"] = usecase["targetDetails"]["include"]
			targetDetails[target]["targetItemIdAttribute"] = usecase["targetDetails"]["targetItemIdAttribute"]
			targetDetails[target]["targetItemNameAttribute"] = usecase["targetDetails"]["targetItemNameAttribute"]
			targetDetails[target]["targetItemTypeAttribute"] = usecase["targetDetails"]["targetItemTypeAttribute"]

			if (usecase['targetDetails'].get('correlation') != None):
				targetDetails[target]["correlation"] = usecase["targetDetails"]["correlation"]

			AllTargets.update(targetDetails)

		if findFinalMappingBySourceaAndUsecase == False:
			mappingItem["targetDetails"].update(AllTargets)
			AllMappings.append(mappingItem)

relCount = 0
attrCount = 0
attrNames = []
for mapItem in AllMappings[0]["map"]:
	# if len(mapItem["targets"])>2:
	# print(mapItem["name"])
	if mapItem["elementType"] == "attribute":
		attrCount = attrCount + 1
		attrNames.append(mapItem["name"])
	if mapItem["elementType"] == "relationship":
		relCount = relCount + 1

print("attr")
print(attrCount)
print("rel")
print(relCount)

for mapItem in AllMappings:
	fileName = output_directory + '/map_' + mapItem['source'] + '_' + mapItem['usecase'] + '.json'

	with open(fileName, 'w') as outfile:
		json.dump(mapItem, outfile, indent=4, sort_keys=True, separators=(',', ': '))

with open(output_directory + "/AllMappings.txt", 'w') as outfile:
	json.dump(AllMappings, outfile, indent=4, sort_keys=True, separators=(',', ': '))
result["maps"] = AllMappings

with open(output_directory + "/raw_maps.txt", 'w') as outfile:
	json.dump(mapping_json, outfile, indent=4, sort_keys=True, separators=(',', ': '))

# print(AllMappings)
# print(json.dumps(AllMappings,indent=4, sort_keys=True))
