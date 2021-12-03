# AutoReport

## Require
Mac os

## Features
- Auto KakaoTalk Send
- auto Daily report generation
- 2 type of excel file auto generation
- Auto generate readme.txt
- save Classes and Students info with json

## How to Use
1. `git clone https://github.com/warpdev/AutoReport.git`
2. edit the path in `PP_Main.py` to your report file folder
3. create `savedata.json` and `path.json`, then fill the content of them with `{}`
4. run!

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
		"excelCol": D,
		"classHomework": "풀던 문제 마무리해서 풀어오기",
		"noHomework": false
	}
}
```
