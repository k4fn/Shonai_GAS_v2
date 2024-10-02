//----------適宜変更してください-----------------------------------------------------------------

//折衝表のスプレッドシートのurl
const sesshoUrl = 'https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxx';

//サークル名前
const circle = 'xxx部';

//作成したい学内集会願の名前(月は後で追加されます)
const newSSfileName = circle + "_学内集会願コピペ_";

//学内集会願のテンプレのurl
const tmpUrl = 'https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxx';

//作成した学内集会願を保存したいフォルダーのid
const id = 'xxxxxxxxxxxxxxx';

//部屋一覧
const room = ['4203', '4302', '5405A', '5405B', '5505', '防音室B', '防音室C'];

//授業開始時間
const startTime = ['9:20', '11:10', '12:50', '13:40', '15:30', '17:20', '19:10', '21:00'];

//授業終了時間
const endTime = ['11:00', '12:50', '13:40', '15:20', '17:10', '19:00', '20:50', '22:00'];

//人数
const users = '5';

//部屋の時間とかを記入するところの行数
const gyou = 14;

//------------------------------------------------------------------------------------------


function myFunction() {
  //折衝表のスプレッドシートの読み込み
  const sesshoSS = SpreadsheetApp.openByUrl(sesshoUrl);

  //学内集会願のスプレッドシートの作成
  const newSS = SpreadsheetApp.create('あ');

  Logger.log(newSS.getUrl());

  //設定した保存したいフォルダーに移動
  if (id != '') {
    const file = DriveApp.getFileById(newSS.getId());
    file.moveTo(DriveApp.getFolderById(id));
  }


  //テンプレを読み込む
  const tmpSheet = SpreadsheetApp.openByUrl(tmpUrl).getSheetByName('テンプレ');

  //部員に周知する用
  let schedule = Array(32);//0-31の32個

  //circleSheetに追加していくとき用の値
  //let setRow = 1;

  //配列roomに入っている部屋すべて繰り返す
  for (let i = 0; i < room.length; i++) {

    let roomRow = 0;

    //部屋のシートを選択
    const roomSheet = sesshoSS.getSheetByName(room[i]);

    //学内集会願の部屋のシートを作成
    let copySheet = tmpSheet.copyTo(newSS);
    copySheet.setName(room[i] + '_1');

    // if(i == 0){
    //   newSS.deleteSheet(newSS.getSheetByName('シート1'));
    // }


    for (let row = 2; row <= 32; row++) {//行に関して、x月1日(2行目)からx月31日(32行目)まで繰り返す

      //時間割の部分を取得
      const range = roomSheet.getRange(row, 2, 1, 8);

      //月、日、曜日を取得
      const date = new Date(roomSheet.getRange(row, 1, 1, 1).getValue());
      const month = date.getMonth() + 1;  //月
      const day = date.getDate(); //日
      const dayOfWeek = ["日", "月", "火", "水", "木", "金", "土"][date.getDay()];  // 曜日(日本語表記)

      //作成したい学内集会願のファイル名に「◯月」を追加
      if (row == 2) {
        newSS.rename(newSSfileName + month + "月");
      }

      //折衝表の時間割の配列
      let timetableArray = [0, 0, 0, 0, 0, 0, 0, 0];//1時限目のをtimetablePerDay[0]

      //利用時間を配列timetableArrayに記録
      for (let col = 2; col <= 9; col++) {//列に関して、１限目(2列目)から放課後(9列目)まで繰り返す

        const a1 = R1C1toA1("R" + row + "C" + col);
        const value = roomSheet.getRange(a1).getValue();

        if (value == circle) {
          timetableArray[col - 2] = 1;
        }
      }
      const mergedRanges = range.getMergedRanges();
      for (var x = 0; x < mergedRanges.length; x++) {
        if (mergedRanges[x].getDisplayValue() == circle) {
          let start = A1Notation_to_start(mergedRanges[x].getA1Notation())
          const end = A1Notation_to_end(mergedRanges[x].getA1Notation())
          for (; start <= end; start++) {
            timetableArray[start - 2] = 1;
          }
        }
      }

      //配列timetableArrayから、具体的な利用時間を取得　ex(9:20~22:00
      let count = 0;
      for (let k = 0; k < timetableArray.length; k++) {
        let start, end;

        if (timetableArray[k] == 1) {
          //roomRow++;
          count++;

          start = k; //k=0は1時限目
          while (timetableArray[k] == 1) {
            k++;
          }
          k--;
          end = k;

          //set = [部屋, 開始時間, 終了時間]　の配列
          let set = [room[i], startTime[start], endTime[end], '(' + dayOfWeek + ')'];

          //ここで学内集会願の表をつくる
          let values = Array(27);
          values[0] = room[i];
          values[5] = month + '/' + day;
          values[6] = '(' + dayOfWeek + ')';
          for (let x = 0; x < timetableArray.length; x++) {
            if (timetableArray[x] == 1) {
              values[6 + (x + 1) * 2 - 1] = '〇';
            }
          }
          values[23] = set[1] + '\n～\n' + set[2];
          values[26] = users;
          values = [values];


          if (count >= 2) {
            let text;
            if (roomRow % gyou == 0) {
              text = copySheet.getRange("AC" + (roomRow)).getValue();
              copySheet.getRange("AC" + (roomRow)).setValue(text + set[1] + '～' + set[2]);
            } else {
              text = copySheet.getRange("AC" + (roomRow % gyou)).getValue();
              copySheet.getRange("AC" + (roomRow % gyou)).setValue(text + set[1] + '～' + set[2]);
            }
          } else {
            roomRow++;
            if (roomRow == gyou + 1) {
              copySheet = tmpSheet.copyTo(newSS);
              copySheet.setName(room[i] + '_' + (Math.floor(roomRow / gyou) + 1));
            }

            //書き込む
            if (roomRow % gyou == 0) {
              copySheet.getRange(R1C1toA1('R' + roomRow + 'C1') + ':' + R1C1toA1('R' + roomRow + 'C27')).setValues(values);
            } else {
              copySheet.getRange(R1C1toA1('R' + (roomRow % gyou) + 'C1') + ':' + R1C1toA1('R' + (roomRow % gyou) + 'C27')).setValues(values);
            }
          }




          //配列scheduleに、日にちごとに配列setを格納
          if (schedule[day] == null) {
            const arr = [set];
            schedule[day] = arr;
          } else {
            schedule[day].push(set);
          }
        }
      }
      if (row == 2) {
        schedule[0] = month;
      }
    }
    if (roomRow == 0) {
      //newSS.deleteSheet(newSS.getSheetByName(room[i]));
      newSS.deleteSheet(copySheet);
    }
  }
  newSS.deleteSheet(newSS.getSheetByName('シート1'));

  //部員に知らせるようのやつ
  let scheduleStr = '';
  for (let day = 1; day <= 31; day++) {
    if (schedule[day] != null) {
      for (let x = 0; x < schedule[day].length; x++) {
        if (x != 0 && schedule[day][x - 1][0] == schedule[day][x][0]) {
          scheduleStr = scheduleStr + ' ' + schedule[day][x][1] + '～' + schedule[day][x][2];
        } else if (x == 0) {
          scheduleStr = scheduleStr + '\n' + schedule[0] + '/' + day + schedule[day][x][3] + ' ' + schedule[day][x][0] + ' ' + schedule[day][x][1] + '～' + schedule[day][x][2];
        } else {
          scheduleStr = scheduleStr + '\n';
          if (day <= 9) {
            scheduleStr = scheduleStr + '             ';//space x 
          } else {
            scheduleStr = scheduleStr + '               ';//space x 
          }
          scheduleStr = scheduleStr + schedule[day][x][0] + ' ' + schedule[day][x][1] + '～' + schedule[day][x][2];
        }
      }
      scheduleStr = scheduleStr + '\n';
    }
  }
  scheduleStr = scheduleStr.substring(1);
  //Logger.log(scheduleStr);

  let scheduleSheet = newSS.insertSheet();
  newSS.moveActiveSheet(newSS.getNumSheets());
  scheduleSheet.setName("周知用");
  scheduleSheet.getRange('A1').setValue(scheduleStr);
}





function R1C1toA1(r1c1) {
  var row = r1c1.match(/R(\d+)/)[1];
  var col = r1c1.match(/C(\d+)/)[1];
  return columnToLetter(col) + row;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function A1toCol(a1) {
  const code = a1.replace(/\d/g, "");
  const col = code.charCodeAt(0) - 64;
  //Logger.log(col);
  return Number(col);
}

function A1Notation_to_start(A1Notation) {
  const index = A1Notation.indexOf(':');
  const a1 = A1Notation.substring(0, index);
  const start = A1toCol(a1);
  return start;
}

function A1Notation_to_start(A1Notation) {
  const index = A1Notation.indexOf(':');
  const a1 = A1Notation.substring(0, index);
  const start = A1toCol(a1);
  return start;
}

function A1Notation_to_end(A1Notation) {
  const index = A1Notation.indexOf(':');
  const a1 = A1Notation.substring(index + 1);
  const end = A1toCol(a1);
  return end;
}
