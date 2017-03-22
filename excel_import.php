

<html>
<head>
<title>成大體育室運動安全問卷</title>
<link rel="stylesheet" href="encryptor.css"> 
</head><center>
<body>

        <div class="menu">
            <div class="content">
                <div class="spacer"></div>
                <div class="item">
                    <div class="title"> <font size="+6"><b>運動安全問卷</b></font></div>
                    <div class="body">
                        <font size="+1">
                            <p>本問卷的目的是在了解您的健康狀況，以增加體適能活動的安全性，本問卷參考美國運動醫學會(1986)之Physical Activity Readiness Questionnaire (PAR-Q)，修訂後使用。請您在參與校內各項體育運動檢測、3000公尺跑步測驗或其他體能性活動前，先回答下列8題。
                          這份問卷會告訴您是否應在開始運動前諮詢醫生。請仔細閱讀下列問題，並誠實答覆：</p>
                        </font>
                    </div>
                    </div>
                    </div>
                    </div>
<b><font size="+1"><p> 請答「是」或「否」</center></p></font></b>
<div class="main">
<form action="excel_import.php" method="post">
　輸入學號驗證&nbsp;:&nbsp; <input type="test" name="YourName"><BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;是&nbsp;&nbsp;&nbsp;否<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q1 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q1 Value="No">1. 醫生曾說過你的心臟有問題，以及您只可進行醫生建議的體能活動？
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q2 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q2 Value="No">2. 您進行體能活動時會感到胸口痛？
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q3 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q3 Value="No">3. 過去一個月內，您曾否在沒有進行體能活動時也感到胸口痛？
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q4 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q4 Value="No">4. 您曾否因感到暈眩而失去平衡，或曾否失去知覺？
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q5 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q5 Value="No">5. 您的骨骼或關節（例如：背、膝或髖）是否有傷病史？或會因改變體能活動而惡化？
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q6 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q6 Value="No">6. 醫生目前是否有開血壓或心臟藥物（例如：利尿劑）給您服用？
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q7 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q7 Value="No">7. 您是否有脊椎側彎？
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT Type=Radio Name=Q8 Value="yes">&nbsp;&nbsp;<INPUT Type=Radio Name=Q8 Value="No">8. 是否有其他理由令您不應進行體能活動？

<BR><center><input type="submit" value="送出表單"></center>
　
</form>
</div>
<div class="display">
<?php
	//引入函式庫
	include 'Classes/PHPExcel.php';
	include 'Classes/PHPExcel/Writer/Excel5.php';
	header("Content-Type:text/html; charset=utf-8");
	//設定要被讀取的檔案，經過測試檔名不可使用中文
	$file = 'load/standardtable.xls';
	try {
	    $objPHPExcel = PHPExcel_IOFactory::load($file);

	} catch(Exception $e) {
	    die('Error loading file "'.pathinfo($file,PATHINFO_BASENAME).'": '.$e->getMessage());
	}
	
	$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
		$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);  //調用Excel 2003以下的版本
	$sheetDataWriter = $objPHPExcel->setActiveSheetIndex(0);



	//學號和問卷參數
$YourName=$_POST[YourName];
$Q1=$_POST[Q1];
$Q2=$_POST[Q2];
$Q3=$_POST[Q3];
$Q4=$_POST[Q4];
$Q5=$_POST[Q5];
$Q6=$_POST[Q6];
$Q7=$_POST[Q7];
$Q8=$_POST[Q8];
date_default_timezone_set('Asia/Taipei');  
$time = date("Y-m-d H:i:s"); 
$tablekey;

//驗證學號
if($YourName == ""){

		echo "您沒有輸入學號";
							echo '<br><br>'; 
}
if($YourName != "" && $YourName != "0"){
//　echo '接收到的內容為:'.$YourName;
	//echo "<h2>列印每一行的資料</h2>";
	//欄與列的index
	$colindex=4;
	$rowindex=0;
	//某行完全沒有值的判斷變數
	$rownull=true;
	//資料對應的欄位標題，有時標題也有利用的空間，完整版我是有用到。
	$title = array();
	foreach($sheetData as $key => $col){
		//讀取標題
		if($rowindex == 0){
			foreach ($col as $colkey => $colvalue){
		//		array_push($title,$colvalue);
			}
		}
		//前面1行不讀入,可更改值設定前幾行不讀取
		if($rowindex>= 1){
			//echo "行{$key}: "."<br>";
			$temp="";
			foreach ($col as $colkey => $colvalue){
			
				//#--為後面使用字串切割的key
				//為第二列資料並且不為最後一列資料，增加切割時用的字串(可更改為不常使用的符號或文字)。輸出格式會變成: 資料1#--資料2#--資料3#--資料4....資料n，每筆資料中間會有#--
				if($colindex > 0 && $colindex != sizeof($col)-1)
				$temp.="#--";
				//前面0列不讀入,可更改值，設定前幾列不讀取，或是為n+1列開始讀取。
				if($colindex == 8){
					//某列值不為空，判斷該行就算有資料。
					if($colvalue!="")
						$rownull=false;
					//將資料暫存下來，繼續讀取下一列。
					$temp.=$colvalue;
				}
				//列的index遞增
				$colindex++;
			}
			//如果設定保護工作表會讀取整份文件Excel,所以$rownull來判斷讀取到的某一行是否完全沒有輸入值
			if($rownull)
				echo "行".$key."沒有值<br>";
			if($rownull && $rowindex > 0)//如果某行完全沒有值，並且讀取到的是內容(標題為第一行,$rowindex=0)，就不在繼續讀取，節省資源。
				break;
			//某行完全沒有值的"判斷變數"，改回預設值
			$rownull=true;
		
			$text=explode("#--",$temp);
			//某一行的所有資料
			for($i=0;$i<sizeof($text);$i++){
				//某列資料值為空
				if($text[$i] == $YourName ){
					$tablekey = $rowindex;
					echo $text[$i]." 學號存在!!";
					echo '<br><br>';  
					$YourName = "0";
		break;
	}
				

			}
			//列的index歸零
			$colindex=4;
			//輸出換行
			//echo "<br/>";
		}
		//if($rowindex>=1)
		//	echo "<hr>";
		//行的index遞增
		$rowindex++;

	} if($YourName != "0"){
		echo $YourName." 學號輸入錯誤";
							echo '<br><br>'; 
	}
}  
/*$M="M".$tablekey;
$N="N".$tablekey;
$O="O".$tablekey;
$P="P".$tablekey;
$Q="Q".$tablekey;
$R="R".$tablekey;
$S="S".$tablekey;
$T="T".$tablekey;*/
$tablekey ++;
//$U="U".$tablekey;

if ($Q1 != "" && $Q2 != "" && $Q3 != "" && $Q4 != "" && $Q5 != "" && $Q6 != "" && $Q7 != "" && $Q8 != ""){

$sheetDataWriter->setCellValue("M".$tablekey,$Q1); 
$sheetDataWriter->setCellValue("N".$tablekey,$Q2); 
$sheetDataWriter->setCellValue("O".$tablekey,$Q3); 
$sheetDataWriter->setCellValue("P".$tablekey,$Q4);
$sheetDataWriter->setCellValue("Q".$tablekey,$Q5); 
$sheetDataWriter->setCellValue("R".$tablekey,$Q6); 
$sheetDataWriter->setCellValue("S".$tablekey,$Q7); 
$sheetDataWriter->setCellValue("T".$tablekey,$Q8);
$sheetDataWriter->setCellValue("U".$tablekey,$time);
//	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007'); 
$objWriter->save('load/standardtable.xls');
					echo " 資料表單輸入完全!!";
}else{
echo "表單輸入不完全!<br><br>";
echo "請輸入表單!!";
}



?>
</div>

<div class = "main2"> <div class="content">
                <div class="spacer"></div>
                <div class="item">
                                <center>    <div class="title"> <font size="+2"><b>一題或以上答「是」</b></font></div> </center>
                    <div class="body">
                        <font size="+1">
                            <p>在參與檢測及活動前，請先跟體育老師或親自與醫生商談，告知醫生這份問卷有疑慮之狀況，以及您回答「是」的問題</p>
                            <ul> 
  <li>如果您想進行練習檢測項目，請先在開始時慢慢進行，然後逐漸增加體能活動量或您只進行一些安全的活動，及告訴醫生您希望參加的活動並聽從他的意見。
  <li>找出一些安全及有益健康的體能活動。 
</ul></font>
<p><i><b>請注意：如因健康狀況轉變，致使您隨後須回答「是」的話，應立即告知醫生或體育老師，看看應否更改您的體能活動。</b></i></p>
                        </font>
                    </div>
                    </div>
                    </div> </div>
<div class = "main3"> <div class="content">
                <div class="spacer"></div>
                <div class="item">
                              <center>      <div class="title"> <font size="+2"><b>全部答「否」</b></font></div> </center>
                    <div class="body">
                        <font size="+1">
                            <p>如果您對這份問卷的全部誠實回答「否」，您有理由確信您可以：</p>
                            <ul>
                            <li>開始增加運動量：開始慢慢進行，然後逐漸增加，這是最安全和最容易的方法。
                            <li>參加體能評估：這是一種確定您基本體能的好方法。此外，亦主張您量血壓；如果讀數超過144/94，請先徵詢醫生的意見。
                            </ul>
                        </font>
                    </div>
                    </div>
                    </div> </div>


<div class = "main4"> <div class="content">
                <div class="spacer"></div>
                <div class="item">
                                 <center>  <div class="title"> <font size="+2"><b>請延遲增加運動量建議</b></font></div> </center>
                    <div class="body">
                        <font size="+1">
                            <p>若有以下情形需先徵詢體育老師或醫生的意見，然後才決定是否增加運度量</p>
                            <ul>
                            <li>如果您因感冒或發燒等暫時性疾病而感到不適，請在康復後才增加運動量。
                            <li>如果您懷孕或可能懷孕，請詢問醫生或體育老師。
                            </ul>
                        </font>
                    </div>
                    </div>
                    </div> </div>

<div class = "Bottom"> <div class="content">
                <div class="spacer"></div>
                <div class="item">
                    <div class="body">
                        <font size="+0.5">
                            <p>備註：如填妥此問卷後有疑問或有未盡之事宜，請先徵詢體育老師或醫生的意見，然後才進行體能活動</p>
                        </font>
                    </div>
                    </div>
                    </div> </div>
</body>
</html>


<?php
/*	//引入函式庫
	include 'Classes/PHPExcel.php';
	header("Content-Type:text/html; charset=utf-8");
	//設定要被讀取的檔案，經過測試檔名不可使用中文
	$file = '103-1.xls';
	try {
	    $objPHPExcel = PHPExcel_IOFactory::load($file);
	} catch(Exception $e) {
	    die('Error loading file "'.pathinfo($file,PATHINFO_BASENAME).'": '.$e->getMessage());
	}
	
	$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
	



	//驗證學號
$YourName=$_POST[YourName];

if($YourName != "" && $YourName != "0"){
//　echo '接收到的內容為:'.$YourName;
	//echo "<h2>列印每一行的資料</h2>";
	//欄與列的index
	$colindex=4;
	$rowindex=0;
	//某行完全沒有值的判斷變數
	$rownull=true;
	//資料對應的欄位標題，有時標題也有利用的空間，完整版我是有用到。
	$title = array();
	foreach($sheetData as $key => $col){
		//讀取標題
		if($rowindex == 0){
			foreach ($col as $colkey => $colvalue){
		//		array_push($title,$colvalue);
			}
		}
		//前面1行不讀入,可更改值設定前幾行不讀取
		if($rowindex>= 1){
			//echo "行{$key}: "."<br>";
			$temp="";
			foreach ($col as $colkey => $colvalue){
			
				//#--為後面使用字串切割的key
				//為第二列資料並且不為最後一列資料，增加切割時用的字串(可更改為不常使用的符號或文字)。輸出格式會變成: 資料1#--資料2#--資料3#--資料4....資料n，每筆資料中間會有#--
				if($colindex > 0 && $colindex != sizeof($col)-1)
				$temp.="#--";
				//前面0列不讀入,可更改值，設定前幾列不讀取，或是為n+1列開始讀取。
				if($colindex == 8){
					//某列值不為空，判斷該行就算有資料。
					if($colvalue!="")
						$rownull=false;
					//將資料暫存下來，繼續讀取下一列。
					$temp.=$colvalue;
				}
				//列的index遞增
				$colindex++;
			}
			//如果設定保護工作表會讀取整份文件Excel,所以$rownull來判斷讀取到的某一行是否完全沒有輸入值
			if($rownull)
				echo "行".$key."沒有值<br>";
			if($rownull && $rowindex > 0)//如果某行完全沒有值，並且讀取到的是內容(標題為第一行,$rowindex=0)，就不在繼續讀取，節省資源。
				break;
			//某行完全沒有值的"判斷變數"，改回預設值
			$rownull=true;
		
			$text=explode("#--",$temp);
			//某一行的所有資料
			for($i=0;$i<sizeof($text);$i++){
				//某列資料值為空
				if($text[$i] == $YourName ){
					echo $text[$i]." 學號存在!!";
					echo " 資料表單輸入完全!!";
					$YourName = "0";
		break;
	}
				

			}
			//列的index歸零
			$colindex=4;
			//輸出換行
			//echo "<br/>";
		}
		//if($rowindex>=1)
		//	echo "<hr>";
		//行的index遞增
		$rowindex++;

	} if($YourName != "0"){
		echo $YourName." 學號輸入錯誤";
	}
} else {
		echo "您沒有輸入學號";
	}*/
?>
