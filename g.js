/** @OnlyCurrentDoc */
function cnum(n){
	n=n-1;
	let alphabet = 'abcdefghijklmnopqrstuvwxyz'.toUpperCase().split('');
	if(n<alphabet.length) return alphabet[n];
	else {
		return alphabet[Math.floor(n/alphabet.length)-1]+""+alphabet[n%alphabet.length];
	}
}

function getPrice(price) {
	return parseFloat((''+price).replaceAll(/[^0-9\.\,]/gim,'').replace(',','.'));
}

function checkData() {
  //var ui = SpreadsheetApp.getUi();
  let d = new Date();
  const month1 = d.getMonth();
	const year1 = d.getFullYear();
  let dt = new Date();
  dt.setMonth(dt.getMonth() - 1);
  const monthAgo = dt.getMonth();
	const yearAgo = dt.getFullYear();
  console.log("This month: ","GL Spend "+ (+month1+1).toString()+" "+year1.toString());
  let sheetNowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GL Spend "+ (+month1+1).toString()+" "+year1.toString());
  console.log("This prev month: ","GL Spend "+ (+monthAgo+1).toString()+" "+yearAgo.toString());
  let monthAgoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GL Spend "+ (+monthAgo+1).toString()+" "+yearAgo.toString());

	if(!sheetNowSheet || !monthAgoSheet) {
    console.log("no pages...!")
    return;
  }

  let nowSheet = sheetNowSheet.getRange("2:2").getValues();
  let oldSheet = monthAgoSheet.getRange("2:2").getValues();

  //let intersection = nowSheet[0].filter(value => oldSheet[0].includes(value));
    //let intersection = nowSheet[0].map(value => (value!='' && oldSheet[0].includes(value))?value+" _ "+oldSheet[0].indexOf(value)+" _ "+cnum(oldSheet[0].indexOf(value)+1):null );
    let intersection = nowSheet[0].map(value => (value!='' && oldSheet[0].includes(value))?cnum(oldSheet[0].indexOf(value)+1):null );
  intersection = intersection.filter(r=>r && r!='A');
 // console.log("new",nowSheet[0])
 // console.log("old",oldSheet[0])

  //console.log("new",nowSheet.join(','))
  //console.log("old",oldSheet.join(','))

//+":"+rg+34
  intersection = intersection.map(rg=>rg+34);
  console.log("RANGES  ",intersection);

  var rangeList  = monthAgoSheet.getRangeList(intersection).getRanges().map(range => range.getValues()[0][0]);
  let total = rangeList.reduce((partialSum, a) => getPrice(partialSum) + a, 0)
  console.log("RESULT ",rangeList);

   console.log("TOTAL  ",total);
   sheetNowSheet.getRange("E43").setValue(getPrice(total).toFixed(2))
}
checkData();
