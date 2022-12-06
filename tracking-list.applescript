# set variables for script
set searchString to "CHICKEN"
set searchString2 to "EGGS"
set searchString3 to "FISH"
set searchString4 to "GROUND PORK"
set searchString5 to "HALAL CHICKEN"
set searchString6 to "MILK"
set foundResult to false

# generalized command, will target visible sheet
tell application "Microsoft Excel"
	insert into range range "E:E" shift shift to right
	tell cell "E10" to set value to "x1.5"
	insert into range range "F:F" shift shift to right
	tell cell "F10" to set value to "cv"
	insert into range range "G:G" shift shift to right
	tell cell "G10" to set value to "usage"
	# loop from bottom, deleting when searchStr found
	tell active workbook
		tell worksheet "Hamper Items  Amount ordered su"
			tell used range
				repeat with i from (count rows) to 11 by -1
					set rowVals to row i's value
					if searchString is in rowVals's item 1 then set foundResult to true
					if searchString2 is in rowVals's item 1 then set foundResult to true
					if searchString3 is in rowVals's item 1 then set foundResult to true
					if searchString4 is in rowVals's item 1 then set foundResult to true
					if searchString5 is in rowVals's item 1 then set foundResult to true
					if searchString6 is in rowVals's item 1 then set foundResult to true
					if foundResult is false then delete row i
					set foundResult to false
				end repeat
			end tell
		end tell
	end tell
	# hardcode calculations
	tell cell "E11" to set value to "=D11*1.5"
	tell cell "F11" to set value to "20"
	tell cell "G11" to set value to "=(E11+f11)/5"
	tell cell "E12" to set value to "=D12*1.5"
	tell cell "F12" to set value to "37"
	tell cell "G12" to set value to "=(E12+f12)/2"
	tell cell "E13" to set value to "=D13*1.5"
	tell cell "F13" to set value to "6"
	tell cell "G13" to set value to "=(E13+f13)/4"
	tell cell "E14" to set value to "=D14*1.5"
	tell cell "F14" to set value to "6"
	tell cell "G14" to set value to "=(E14+f14)"
	tell cell "E15" to set value to "=D15*1.5"
	tell cell "F15" to set value to "20"
	tell cell "G15" to set value to "=(E15+f15)/5"
	tell cell "E16" to set value to "=D16*1.5"
	tell cell "F16" to set value to "60"
	tell cell "G16" to set value to "=(E16+f16)/12"
	display dialog "Operation Complete"
end tell