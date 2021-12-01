# AutoReport

## Features
Auto KakaoTalk Send, auto Daily report generation, 2 type of excel file auto generation, save Classes and Students info with json.

## How to Use
1. clone
2. edit the path in PP_Main.py to your report file folder
3. run!

## savedata.json example
```json
{
	"classname/day-10:00/who": {
		"students": {
			"name1": {
				"isHome": true,
				"feedback": "",
				"kakaoName": "kakaoid",
				"noKakao": false
			},
			"name2": {
				"isHome": false,
				"feedback": "",
				"kakaoName": "",
				"noKakao": false
			}
		},
		"folderName": "10시_classname_who",
		"classComment": "",
		"classSpecial": "",
		"excelCol": 4
	}
}
```