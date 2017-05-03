sample code:

	require('strext')
	local xlsx = require('xlsxio')

	local r = xlsx.read('d:/1.xlsx', { totable = true })

	for i = 1, #r do
		local sheet = r[i]
		print('sheet', i, utf8str.togbk(sheet.name))
		
		for k = 1, #sheet.rows do
			local row = sheet.rows[k]
			print(row[1], row[2], row[3], row[4])
		end
	end