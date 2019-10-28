<?php
// Log Viewer pulls last 100 Lookups from the database
// Total tax rate contains State Rate, County Rate and applicable municipal sales tax rate
// when the location is within a municipal boundary determined by the locator x and y


// Initialize connection variables 
// Database host
$dbserver = 'databasehost';
$dbname = "databasename";
$dbuser = "databaseuser";
$dbpass = "databasepass";

// Web host
$webserver = 'http://myhost.com/SSTLookup/admin/';

// ArcGIS Services
$complocator = 'http://gis.arkansas.gov/arcgis/rest/services/Locator/ASDI_Composite_Locator/GeocodeServer';
$locator1 = 'APF_LOCATOR';
$locator2 = 'ACF_LOCATOR';
$locator3 = 'ZIP9_LOCATOR';
$locator4 = 'ZIP_PLUS_4_LOC';
$boundaries1 = 'http://gis.arkansas.gov/arcgis/rest/services/FEATURESERVICES/Boundaries/MapServer';

//GeoScore of 92 for best results
//scores less than 92 will return results with mis-matched pre/post directionals or street types
$geoscore = 92;

//Log count controls the number of records returned. numbers greater than 100 may be slow to load -- todo paiginated 
$logcount = 100;
 
// Initialize handling variables
$Ziponly = 0;
$Addonly = 0;
$html = "";
$qry = 0;
$qry = $_GET['qry'];
$poss_results = '';

//Database hooks
$connectionInfo = array( "Database"=>$dbname, "UID"=>$dbuser, "PWD"=>$dbpass);
$conn = sqlsrv_connect( $dbserver, $connectionInfo);
// OUTPUT HEADER	
?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Sales and Use Tax Section Local Tax Lookup Tools : Administration : Log Viewer</title>
</head>

<body>
<form id="form1" runat="server">
<?php 

// Is this data coming from a query string or a form post
if ($qry){
	$iStreet = $_GET['Street'];
	$iCity = $_GET['City'];
	$iZIP = $_GET['ZIP'];
	$iZIP5 = $_GET['ZIP5'];
	$iZIP4 = $_GET['ZIP4'];
	if( $conn ) {
		//echo "Connection established.<br />";
		// Get the Default rates for our Jurisdiction
		$sql = "select * from tblSSTP_RATE WHERE JURISDICTION_TYPE = '45'";
		$stmt = sqlsrv_query( $conn, $sql);
		$row = sqlsrv_fetch_array( $stmt );
		$GENRATE = $row['GEN_TAX_RATE_INTRASTATE'] * 100;
		$GENRATEst = $GENRATE . "%";
		$FOODRATE = $row['FOODDRUG_TAX_RATE_INTRA'] * 100;
		$FOODRATEst = $FOODRATE . "%";
		$MFGRATE = $row['MFG_UTILITY_TAX_RATE'] * 100;
		$MFGRATEst = $MFGRATE . "%";
		sqlsrv_free_stmt( $stmt );	 
	}else{
		echo "Connection could not be established the SQL Server.<br />";
		die( print_r( sqlsrv_errors(), true));
	}

	// Are we doing Zip Lookup for Rates or  via the Address Geocode
	// First we check that the Zip5 field ,from the form, has a value
	// If it doesn't, then we move on to the next part of the conditional statement

	// ZIP5 Lookup check
	if ($iZIP5 != ''){
		if ($iZIP4 == ''){
			// we can't have an empty string
			// Need error code page
			echo "Error no Zip4";
			exit(0);
		}

		// Set the ziponly value for later on in the code to distinguish the output of the script as Zip or Address
		$Ziponly = 1;

		// Pull in the posted values and clean them up
		$Zip5 = $iZIP5;
		$Zip4 = $iZIP4;
		$Zip9 = trim($Zip5) . trim($Zip4);
		// Zip9b used for creating queries that require Zip with dash '-'
		$Zip9b = trim($Zip5) . '-' . trim($Zip4);
		
		// Pull the associated Tax rate and jurisdiction codes based on the $Zip9 value
		$sql = "SELECT DISTINCT tblSSTP_BOUNDARY.CITY_NAME, tblSSTP_BOUNDARY.COUNTY_NAME, tblSSTP_RATE.GEN_TAX_RATE_INTRASTATE,
		tblSSTP_RATE.JURISDICTION_TYPE, tblSSTP_RATE.JURISDICTION_FIPS, tblSSTP_RATE.DFA_CODE 
		FROM tblSSTP_BOUNDARY INNER JOIN tblSSTP_RATE ON tblSSTP_RATE.JURISDICTION_FIPS = tblSSTP_BOUNDARY.FIPS_COUNTY 
		OR tblSSTP_RATE.JURISDICTION_FIPS = tblSSTP_BOUNDARY.FIPS_PLACE 
		WHERE tblSSTP_BOUNDARY.ZIP_NINE = '".$Zip9."' order by tblSSTP_RATE.JURISDICTION_TYPE asc";

		$stmt = sqlsrv_query( $conn, $sql);
		$zcnt = 0;
		// We should return only one record but we will cycle through the query results 
		while( $row = sqlsrv_fetch_array( $stmt ) ) {
			$CITY[$zcnt] = $row['CITY_NAME'];
			$COUNTY[$zcnt] = $row['COUNTY_NAME'];
			$GENRATE1[$zcnt] = $row['GEN_TAX_RATE_INTRASTATE'] * 100;
			$GENRATE1st[$zcnt] = $GENRATE1[$zcnt] . "%";
			$JURTYPE[$zcnt] = $row['JURISDICTION_TYPE'];
			$JURFIPS[$zcnt] = $row['JURISDICTION_FIPS'];
			$DFACODE[$zcnt] = $row['DFA_CODE'];
			$zcnt++;
		}
		$zcnt = $zcnt - 1; //remove loops extra count
		sqlsrv_free_stmt( $stmt );

		// Let JSON and the Locators output our results
		// Create singleline for ZIP locator
		$singleline = "ZIP=". urlencode($Zip9b);
		$url = $complocator . "/findAddressCandidates?Street=&City=&State=&" . $singleline . "&category=&outFields=*&forStorage=false&f=pjson";

		$response = file_get_contents($url);
		$candidates = json_decode($response);
		$cnt = 0;
		// Only candidate scores of 92 or better accepted
		// All other scores will be found using the ZIP9_Locator for any Address points not currently identified by using the APF_Locator
		// No ACF data in the SQL database to pull ZIP9 for rates so the ACF_Locator won't be used as a final backup to find addresses
		foreach($candidates->candidates as $candidate) {
			if ( $candidate->attributes->Score > $geoscore && $candidate->attributes->Loc_name == $locator4) {
				// Initialize these match temp vars
				$Matches[0] = "";
				$Matches[1] = "";
				$Matches[2] = "";
				$Matches[3] = "";
				$address[$cnt] = $candidate->address;
				$x[$cnt] = $candidate->location->x;
				$y[$cnt] = $candidate->location->y;
				//Which locator found this match
				$Loc_name[$cnt] = $candidate->attributes->Loc_name;
				//If not address match was found, the locator returns possible matches
				//Lets grab the first four if applicable for outputing later
				$Match_addr[$cnt] = $candidate->attributes->Match_addr;
				$Matches = explode(", ", $Match_addr[$cnt]);
				$Matches1[$cnt] = $Matches[0];
				$Matches2[$cnt] = $Matches[1];
				$Matches3[$cnt] = $Matches[2];
				$Matches4[$cnt] = $Matches[3];

				$City[$cnt] = $candidate->attributes->City;
				$ZIP[$cnt] = $candidate->attributes->ZIP;
				//User_fld originally gave other values with Zip lookups but now returns nothing. 
				//we set a value as the original code assumed it would return one
				$User_fld[$cnt] = 777;
				$cnt++;
			}
		}
		$cnt=$cnt-1;
		
		// Locator above returns updates user_fld field to 777 if match found from the locator. 
		// Lets make sure that we have data to work with
		if (is_numeric($User_fld[0])){			
			$url2 = $boundaries1 . "/identify?geometry=".$x[0]."%2C+".$y[0]."&geometryType=esriGeometryPoint&sr=26915&layers=visible%3A+8%2C41&layerDefs=&time=&layerTimeOptions=&tolerance=0&mapExtent=355223.7187%2C+3652712.3124%2C+798032.9088%2C+4041425.8755&imageDisplay=600%2C550&returnGeometry=false&maxAllowableOffset=&geometryPrecision=&dynamicLayers=&returnZ=false&returnM=false&gdbVersion=&f=pjson";
			$response2 = file_get_contents($url2);
			$results = json_decode($response2);
			
			foreach($results->results as $result) {
				if ($result->attributes->FHWA_Numbe != '') {
					$CountyFIPS = $result->attributes->FHWA_Numbe;
				}
				if ($result->attributes->County_Nam != '') {
					$CountyName = $result->attributes->County_Nam;
				}			
				if ($result->attributes->CITY_FIPS != '') {
					$CityFIPS = $result->attributes->CITY_FIPS;
				}
				if ($result->attributes->CITY_NAME != '') {
					$CityName = $result->attributes->CITY_NAME;
				}
			}
			if (strlen($CountyFIPS)<3){
				if (strlen($CountyFIPS)<2){
					$CountyFIPS = "00".$CountyFIPS;
				} else {
					$CountyFIPS = "0".$CountyFIPS;
				}
			}			
			$sql = "SELECT DISTINCT tblSSTP_RATE.GEN_TAX_RATE_INTRASTATE, tblSSTP_RATE.JURISDICTION_TYPE, tblSSTP_RATE.JURISDICTION_FIPS, tblSSTP_RATE.DFA_CODE FROM tblSSTP_RATE WHERE tblSSTP_RATE.JURISDICTION_FIPS = '".$CountyFIPS."' OR tblSSTP_RATE.JURISDICTION_FIPS = '".$CityFIPS."' order by tblSSTP_RATE.JURISDICTION_TYPE asc";			

			$stmt = sqlsrv_query( $conn, $sql);
			$zcnt = 0;
			while( $row = sqlsrv_fetch_array( $stmt ) ) {
				$CITY[$zcnt] = $CityName;
				$COUNTY[$zcnt] = $CountyName;
				$GENRATE1[$zcnt] = $row['GEN_TAX_RATE_INTRASTATE'] * 100;
				$GENRATE1st[$zcnt] = $GENRATE1[$zcnt] . "%";
				$JURTYPE[$zcnt] = $row['JURISDICTION_TYPE'];
				$JURFIPS[$zcnt] = $row['JURISDICTION_FIPS'];
				$DFACODE[$zcnt] = $row['DFA_CODE'];
				$zcnt++;
			}
			$zcnt = $zcnt - 1;
			sqlsrv_free_stmt( $stmt );	
		}
	} else { 
		// Address Lookup if we have No Ziplookup posted input fields
		// Lets geocode this data then after checking for empty values again
		if ($iStreet == '' || $iCity == '' || $iZIP == ''){
			echo "Error. No Address data";
			exit(0);
		}
		// Let JSON and the Locators do all the work
		$singleline = "SingleLine=". urlencode($iStreet. " " . $iCity. " AR" . " " . $iZIP);
		$url = $complocator . "/findAddressCandidates?" . $singleline . "&category=&outFields=*&forStorage=false&f=pjson";

		$response = file_get_contents($url);
		$candidates = json_decode($response);
		$cnt = 0;

		// Only candidate scores of 92 or better accepted : Set $geoscore to change that to another value (not recommended)
		// All other scores will be found using the ZIP9_Locator for any Address points not currently identified by using the APF_Locator 
		foreach($candidates->candidates as $candidate) {
			if ( $candidate->attributes->Score > $geoscore && (($candidate->attributes->Loc_name == $locator1) || ($candidate->attributes->Loc_name == $locator3) || ($candidate->attributes->Loc_name == $locator2))) {
				$Matches[0] = "";
				$Matches[1] = "";
				$Matches[2] = "";
				$Matches[3] = "";
				$address[$cnt] = $candidate->address;
				$x[$cnt] = $candidate->location->x;
				$y[$cnt] = $candidate->location->y;
				$Loc_name[$cnt] = $candidate->attributes->Loc_name;
				$Match_addr[$cnt] = $candidate->attributes->Match_addr;
				$Matches = explode(", ", $Match_addr[$cnt]);
				$Matches1[$cnt] = $Matches[0];
				$Matches2[$cnt] = $Matches[1];
				$Matches3[$cnt] = $Matches[2];
				$Matches4[$cnt] = $Matches[3];

				$City[$cnt] = $candidate->attributes->City;
				$ZIP[$cnt] = $candidate->attributes->ZIP;
				//Zip locator returns no value for user_fld so we set one, 
				//otherwise we assign the value returnd by the other locators
				if ( $candidate->attributes->Loc_name == $locator3 ){
					$User_fld[$cnt] = 777;
				} else {
					$User_fld[$cnt] = $candidate->attributes->User_fld;
				}
				$cnt++;
			}
		}
		$cnt=$cnt-1;
		
		if ($User_fld[0] != '' && is_numeric($User_fld[0])){
			$url2 = $boundaries1 . "/identify?geometry=".$x[0]."%2C+".$y[0]."&geometryType=esriGeometryPoint&sr=26915&layers=visible%3A+8%2C41&layerDefs=&time=&layerTimeOptions=&tolerance=0&mapExtent=355223.7187%2C+3652712.3124%2C+798032.9088%2C+4041425.8755&imageDisplay=600%2C550&returnGeometry=false&maxAllowableOffset=&geometryPrecision=&dynamicLayers=&returnZ=false&returnM=false&gdbVersion=&f=pjson";
			$response2 = file_get_contents($url2);
			$results = json_decode($response2);
			
			foreach($results->results as $result) {
				if ($result->attributes->FHWA_Numbe != '') {
					$CountyFIPS = $result->attributes->FHWA_Numbe;
				}
				if ($result->attributes->County_Nam != '') {
					$CountyName = $result->attributes->County_Nam;
				}			
				if ($result->attributes->CITY_FIPS != '') {
					$CityFIPS = $result->attributes->CITY_FIPS;
				}
				if ($result->attributes->CITY_NAME != '') {
					$CityName = $result->attributes->CITY_NAME;
				}
			}
			if (strlen($CountyFIPS)<3){
				if (strlen($CountyFIPS)<2){
					$CountyFIPS = "00".$CountyFIPS;
				} else {
					$CountyFIPS = "0".$CountyFIPS;
				}
			}			
			
			$sql = "SELECT DISTINCT tblSSTP_RATE.GEN_TAX_RATE_INTRASTATE, tblSSTP_RATE.JURISDICTION_TYPE, tblSSTP_RATE.JURISDICTION_FIPS, tblSSTP_RATE.DFA_CODE FROM tblSSTP_RATE WHERE tblSSTP_RATE.JURISDICTION_FIPS = '".$CountyFIPS."' OR tblSSTP_RATE.JURISDICTION_FIPS = '".$CityFIPS."' order by tblSSTP_RATE.JURISDICTION_TYPE asc";			

			$stmt = sqlsrv_query( $conn, $sql);
			$zcnt = 0;
			while( $row = sqlsrv_fetch_array( $stmt ) ) {
				$CITY[$zcnt] = $CityName;
				$COUNTY[$zcnt] = $CountyName;
				$GENRATE1[$zcnt] = $row['GEN_TAX_RATE_INTRASTATE'] * 100;
				$GENRATE1st[$zcnt] = $GENRATE1[$zcnt] . "%";
				$JURTYPE[$zcnt] = $row['JURISDICTION_TYPE'];
				$JURFIPS[$zcnt] = $row['JURISDICTION_FIPS'];
				$DFACODE[$zcnt] = $row['DFA_CODE'];
				$zcnt++;
			}
			$zcnt = $zcnt - 1;
			sqlsrv_free_stmt( $stmt );	
		}
	}			
	 
	 // Zip Lookup Results
	if ($Ziponly) {
		$html .= "<table>
		<tr>
			<td></td>
			<td align='left' style='font-weight:bold;'>
								<br /><br />
								<font color='#333399' size='4' face='Arial, Helvetica, sans-serif'><em><u>Excise Taxes By Address</u></em></font>
								<br /><br />
			</td>
		</tr><tr>
			<td>
			</td>
			<td></td>
		</tr><tr>
			<td></td>
			<td>
								<br />
			</td>
		</tr><tr>
			<td></td>
			<td align='left'>
				<table cellspacing='0' cellpadding='0' rules='rows' style='border-width:1px;border-style:Inset;width:250px;border-collapse:collapse;'>
					<tr>
						<td style='background-color:#E7F3FF;font-weight:bold;'>Found County and City</td>
					</tr><tr align='left'>
						<td style='font-weight:bold;'>County: ".$COUNTY[0]."</td>
					</tr><tr align='left'>
					<td style='font-weight:bold;'>City: ".$CITY[0]."</td>
					</tr>
				</table>
								<br />
			</td>
		</tr><tr>
			<td></td>
			<td align='left'>
				<table cellspacing='0' cellpadding='0' rules='all' style='border-width:1px;border-style:Inset;border-collapse:collapse;'>
					<tr align='center' style='background-color:#E7F3FF;font-weight:bold;'>
						<td style='width:183px;'>Tax Entity</td>
						<td style='width:84px;'>Name</td>
						<td style='width:46px;'>FIPS</td>
						<td style='width:57px;'>DFA Code</td>
						<td style='width:52px;'>Rate</td>
						<td style='width:52px;'>Reduced Food Rate</td>
						<td style='width:50px;'>Mfg Utility Rate</td>
					</tr><tr style='height:31px;'>
						<td>State</td><td align='center'>Arkansas</td>
						<td align='center'>05</td>
						<td align='center'></td>
						<td align='center'>".$GENRATEst."</td>
						<td align='center'>".$FOODRATEst."</td>
						<td align='center'>".$MFGRATEst."</td>
					</tr><tr style='height:30px;'>
						<td>County</td>
						<td align='center'>".$COUNTY[0]."</td>
						<td align='center'>".$JURFIPS[0]."</td>
						<td align='center'>".$DFACODE[0]."</td>
						<td align='center'>".$GENRATE1st[0]."</td>
						<td align='center'>".$GENRATE1st[0]."</td>
						<td align='center'>".$GENRATE1st[0]."</td>
					</tr><tr style='height:30px;'>
						<td>City</td>
						<td align='center'>".$CITY[1]."</td>
						<td align='center'>".$JURFIPS[1]."</td>
						<td align='center'>".$DFACODE[1]."</td>
						<td align='center'>".$GENRATE1st[1]."</td>
						<td align='center'>".$GENRATE1st[1]."</td>
						<td align='center'>".$GENRATE1st[1]."</td>
					</tr><tr style='height:30px;'>
						<td></td>
						<td></td>
						<td></td>
						<td align='center' style='font-weight:bold;'>Total:</td>
						<td align='center' style='background-color:#E7F3FF;font-weight:bold;'>".($GENRATE + $GENRATE1[0] + $GENRATE1[1])."%</td>
						<td align='center' style='background-color:#E7F3FF;font-weight:bold;'>".($FOODRATE + $GENRATE1[0] + $GENRATE1[1])."%</td>
						<td align='center' style='background-color:#E7F3FF;font-weight:bold;'>".($MFGRATE + $GENRATE1[0] + $GENRATE1[1])."%</td>
					</tr>
				</table><!-- tax table -->
							<br />
			</td>
		</tr>
	</table>";
	}
	//Address Lookup Results
	if ($Loc_name[0] == $locator1) {
		$Addonly = 1;
		$OutZip = $ZIP[0];
	}
	// ZIP9_Locator will find the best results for any bad data sent
	if ($Loc_name[0] == $locator3) {
		$Addonly = 1;
		$OutZip = $ZIP[0];
	}
	// ACF_Locator will find addresses in rare instances before ZIP9 so lets code for it
	if ($Loc_name[0] == $locator2) {
		$Addonly = 1;
		// Zip output used to have dashes. This removed the dash
		//$OutZip = substr_replace($ZIPNINE[0], "-", 5, 0);
		$OutZip = $ZIP[0];
	}
	if ($Addonly){
		$html .= "<table>
		<tr>
			<td></td>
			<td align='left' style='font-weight:bold;'>
								<br /><br />
								<font color='#333399' size='4' face='Arial, Helvetica, sans-serif'><em><u>Excise Taxes By Address</u></em></font>
								<br /><br />
			</td>
		</tr><tr>
			<td>
			</td>
			<td>
				<table style='width:524px;'>
					<tr>
						<td>
							<table cellspacing='0' cellpadding='0' rules='rows' style='border-width:1px;border-style:Inset;height:65px;width:300px;border-collapse:collapse;'>
								<tr style='background-color:#E7F3FF;'>
									<td style='font-weight:bold;'>Input Mailing Address 
									</td>
									<td> 
													<a href='javascript: history.go(-1)'>(Change Input Add.)</a>
									</td>
								</tr><tr align='left'>
									<td colspan='2'>".$iStreet."</td>
								</tr><tr align='left'>
									<td colspan='2'>".$iCity.", ".$iZIP."</td>
								</tr>
							</table>    
							<br />
						</td><td align='left'>
							<table cellspacing='0' cellpadding='0' rules='rows' style='border-width:1px;border-style:Inset;height:65px;width:300px;border-collapse:collapse;'>
								<tr style='background-color:#E7F3FF;'>
									<td style='font-weight:bold;'>Output Tax Jurisdiction Address </td>
								</tr><tr>
									<td>".$Matches1[0]."</td>
								</tr><tr>
									<td>".$Matches2[0].", ".$OutZip."</td>
								</tr>
							</table>    
										<br />
						</td>
					</tr>
				</table>
			</td>
		</tr><tr>
			<td></td>
			<td>
								<br />
			</td>
		</tr><tr>
			<td></td>
			<td align='left'>
				<table cellspacing='0' cellpadding='0' rules='rows' style='border-width:1px;border-style:Inset;width:250px;border-collapse:collapse;'>
					<tr>
						<td style='background-color:#E7F3FF;font-weight:bold;'>Found County and City</td>
					</tr><tr align='left'>
						<td style='font-weight:bold;'>County: ".$COUNTY[0]."</td>
					</tr><tr align='left'>
					<td style='font-weight:bold;'>City: ".$CITY[0]."</td>
					</tr>
				</table>
								<br />
			</td>
		</tr><tr>
			<td></td>
			<td align='left'>
				<table cellspacing='0' cellpadding='0' rules='all' style='border-width:1px;border-style:Inset;border-collapse:collapse;'>
					<tr align='center' style='background-color:#E7F3FF;font-weight:bold;'>
						<td style='width:183px;'>Tax Entity</td>
						<td style='width:84px;'>Name</td>
						<td style='width:46px;'>FIPS</td>
						<td style='width:57px;'>DFA Code</td>
						<td style='width:52px;'>Rate</td>
						<td style='width:52px;'>Reduced Food Rate</td>
						<td style='width:50px;'>Mfg Utility Rate</td>
					</tr><tr style='height:31px;'>
						<td>State</td><td id='TableCell20' align='center'>Arkansas</td>
						<td align='center'>05</td>
						<td align='center'></td>
						<td align='center'>".$GENRATEst."</td>
						<td align='center'>".$FOODRATEst."</td>
						<td align='center'>".$MFGRATEst."</td>
					</tr><tr style='height:30px;'>
						<td>County</td>
						<td align='center'>".$COUNTY[0]."</td>
						<td align='center'>".$JURFIPS[0]."</td>
						<td align='center'>".$DFACODE[0]."</td>
						<td align='center'>".$GENRATE1st[0]."</td>
						<td align='center'>".$GENRATE1st[0]."</td>
						<td align='center'>".$GENRATE1st[0]."</td>
					</tr><tr style='height:30px;'>
						<td>City</td>
						<td align='center'>".$CITY[1]."</td>
						<td align='center'>".$JURFIPS[1]."</td>
						<td align='center'>".$DFACODE[1]."</td>
						<td align='center'>".$GENRATE1st[1]."</td>
						<td align='center'>".$GENRATE1st[1]."</td>
						<td align='center'>".$GENRATE1st[1]."</td>
					</tr><tr style='height:30px;'>
						<td></td>
						<td></td>
						<td></td>
						<td align='center' style='font-weight:bold;'>Total:</td>
						<td align='center' style='background-color:#E7F3FF;font-weight:bold;'>".($GENRATE + $GENRATE1[0] + $GENRATE1[1])."%</td>
						<td align='center' style='background-color:#E7F3FF;font-weight:bold;'>".($FOODRATE + $GENRATE1[0] + $GENRATE1[1])."%</td>
						<td align='center' style='background-color:#E7F3FF;font-weight:bold;'>".($MFGRATE + $GENRATE1[0] + $GENRATE1[1])."%</td>
					</tr>
				</table><!-- tax table -->
							<br />
			</td>
		</tr>
	</table>";
	}

	// Address wasn't found or we have multiple inputs with bad data
	// Lets output the matches we did find and give the user the option to correct their input
	// or pick an address tha more closely matches their search
	if ((!$Ziponly) && (!$Addonly)){
		$response = file_get_contents($url);
		$candidates = json_decode($response);
		$cnt = 0;
		foreach($candidates->candidates as $candidate) {
			// We only oupt the what these three locators found
			if ( $candidate->attributes->Loc_name == $locator1 || $candidate->attributes->Loc_name == $locator3 || $candidate->attributes->Loc_name == $locator2 ) {
				$Matches[0] = "";
				$Matches[1] = "";
				$Matches[2] = "";
				$Matches[3] = "";
				$address[$cnt] = $candidate->address;
				list($address1[$cnt], $city1[$cnt], $state1[$cnt], $zip1[$cnt]) = explode(",",$address[$cnt]);
				$x[$cnt] = $candidate->location->x;
				$y[$cnt] = $candidate->location->y;
				$Loc_name[$cnt] = $candidate->attributes->Loc_name;
				$Match_addr[$cnt] = $candidate->attributes->Match_addr;
				$Matches = explode(", ", $Match_addr[$cnt]);
				$Matches1[$cnt] = $Matches[0];
				$Matches2[$cnt] = $Matches[1];
				$Matches3[$cnt] = $Matches[2];
				$Matches4[$cnt] = $Matches[3];
				$PreDir[$cnt] = $candidate->attributes->PreDir;
				$PreType[$cnt] = $candidate->attributes->PreType;
				$StreetName[$cnt] = $candidate->attributes->StreetName;
				$SufType[$cnt] = $candidate->attributes->SufType;
				$SufDir[$cnt] = $candidate->attributes->SufDir;
				$City[$cnt] = $candidate->attributes->City;
				$ZIP[$cnt] = $candidate->attributes->ZIP;
				$User_fld[$cnt] = $candidate->attributes->User_fld;
				$FromAddr[$cnt] = $candidate->attributes->FromAddr;
				$ToAddr[$cnt] = $candidate->attributes->ToAddr;
				$cnt++;
			}
		}
		$cnt=$cnt-1;
		// We limit the number of alternate matches to the top nine	
		// even though more may have been found
		if ($cnt > 9){
			$cnt = 9;
		}
		// if address[0] has no value then we found nothing
		// let the user know and stop here
		if ($address[0] == ''){
			$html .= "                    <table>
		<tr>
			<td></td><td>
								<br />
							</td>
		</tr><tr align='center'>
			<td></td><td>
										<center><strong><em><u>The location was not found!</u></em></strong><br /><br /></center><br />
								Please verify that the information was entered correctly. <a href='javascript: history.go(-1)'>Click here to go back.</a> <br />
								<br />";
		} else {
			$html .= "                    <table>
		<tr>
			<td></td><td>
								<br />
							</td>
		</tr><tr align='center'>
			<td></td><td>
										<center><strong><em><u>The exact location was not found. Here is a list of possible matches.</u></em></strong><br /><br /></center><br />
								Please select a candidate below or <a href='javascript: history.go(-1)'>click here to go back and reenter the address.</a> <br />
								<br />";
			for ($i=0;$i<=$cnt;$i++) {
				$poss_results .= $address1[$i].",".$city1[$i].",".$zip1[$i].",";
				$address1[$i] = urlencode($address1[$i]);
				$city1[$i] = urlencode($city1[$i]);
				$zip1[$i] = urlencode($zip1[$i]);
				$html .= "<a href='".$webserver."LogViewer.php?qry=1&Street=".$address1[$i]."&City=".$city1[$i]."&ZIP=".$zip1[$i]."'>".$address[$i]."</a><br />";
			}							
		}		
		$html .= "</td>
		</tr>
	</table><!-- page table -->";
	}
	 echo $html;
} 
if( $conn ) {
	//echo "Connection established.<br />";
	// Load log header and data
	echo "<table>
		<tr>
			<td></td>
			<td align='left' style='font-weight:bold;'>
								<br /><br />
								<font color='#333399' size='4' face='Arial, Helvetica, sans-serif'><em><u>Lookup Log Viewer</u></em></font>
								<br /><br />
			</td>
		</tr><tr>
			<td'>
			</td>
			<td>	
			</td>
		</tr><tr>
			<td></td>
			<td>
			</td>
		</tr><tr>
			<td></td>
			<td align='left'>
			</td>
		</tr><tr>
			<td></td>
			<td align='left'>
			<table cellspacing='0' cellpadding='0' rules='all' style='border-width:1px;border-style:Inset;border-collapse:collapse;'>
			<tr align='center' style='background-color:#E7F3FF;'>
				<td>Date and Time</td>
				<td>Input Address</td>
				<td>Input City</td>
				<td>Input Zip</td>
				<td>Input Zip Lookup</td>
				<td>Input Zip4</td>
				<td>Lookup Results</td>
				<td>Output County</td>
				<td>Output City</td>
				<td>Output Zip+4</td>
				<td>Total Rate</td>
				<td>Total RFR</td>
				<td>Total MUR</td>
				<td>Possible Matches (each link follows alternate)</td>						
			</tr>";
	$sql = "select TOP(".$logcount.") CONVERT(nvarchar(20),DateCreated, 100) as DateCreated2,Street ,City ,Zip ,Zip5 ,Zip4 ,Results ,fCounty ,fCity ,iStreet ,iCity ,iZIP ,oStreet ,oCity ,oZIP ,StateRate ,StateRFR ,StateMUR ,DFACode ,CountyFIPS ,CountyRate ,CountyRFR ,CountyMUR ,CityFIPS ,CityDFACode ,CityRate ,CityRFR ,CityMUR ,TotalRate ,TotatRFR ,TotalMUR,PossibleMatches from tblSSTP__LookupLog ORDER BY DateCreated DESC";
	$stmt = sqlsrv_query( $conn, $sql);
	$i = 0;
	while( $row = sqlsrv_fetch_array( $stmt ) ) {
		$address1[$i] = urlencode($row['Street']);
		$city1[$i] = urlencode($row['City']);
		$zip1[$i] = urlencode($row['Zip']);
		$zip5[$i] = urlencode($row['Zip5']);
		$zip4[$i] = urlencode($row['Zip4']);
		$loghref = "<a href='".$webserver."LogViewer.php?qry=1&Street=".$address1[$i]."&City=".$city1[$i]."&ZIP=".$zip1[$i]."&ZIP5=".$zip5[$i]."&ZIP4=".$zip4[$i]."'>";
		$possmatchs = "";
		//echo $row['PossibleMatches'];
		if (strpos($row['PossibleMatches'], ',') > 0){
			$r = explode(',', $row['PossibleMatches']);
			$pcount = count($r);
			//list($address2,$city2,$zip2) = explode(',', $row['PossibleMatches']);
			//echo "pcount:".$pcount."<br>";
			for($z=0;$z<=$pcount-3;$z+=3){
				//echo $z."<br>";
				$atemp = $r[0+$z].", ".$r[1+$z]." ,".$r[2+$z];
				$r[0+$z] = urlencode($r[0+$z]);
				$r[1+$z] = urlencode($r[1+$z]);
				$r[2+$z] = urlencode($r[2+$z]);
				$possmatchs .= "<a href='".$webserver."LogViewer.php?qry=1&Street=".$r[0+$z]."&City=".$r[1+$z]."&ZIP=".$r[2+$z]."'>".$atemp."</a><br>";
			}
		}
		echo "<tr id='TableRow16a' style='height:31px;'>
					<td id='TableCell19'>".$loghref.$row['DateCreated2']."</a></td>
					<td id='TableCell20'>".$loghref.$row['Street']."</a></td>
					<td id='TableCell21'>".$loghref.$row['City']."</a></td>
					<td id='TableCell22'>".$loghref.$row['Zip']."</a></td>
					<td id='tblCStateRate'>".$loghref.$row['Zip5']."</a></td>
					<td id='tblCStateFood'>".$loghref.$row['Zip4']."</a></td>
					<td id='tblCStateMfg'>".$loghref.$row['Results']."</a></td>
					<td id='TableCell19'>".$loghref.$row['fCounty']."</a></td>
					<td id='TableCell20'>".$loghref.$row['fCity']."</a></td>
					<td id='TableCell21'>".$loghref.$row['iZIP']."</a></td>
					<td id='TableCell22'>".$loghref.$row['TotalRate']."</a></td>
					<td id='tblCStateRate'>".$loghref.$row['TotatRFR']."</a></td>
					<td id='tblCStateFood'>".$loghref.$row['TotalMUR']."</a></td>
					<td id='tblCStateMfg'>".$possmatchs."</td>						
				</tr>";
		$i++;
	}
	echo "</table><!-- tax table -->
					<br />
			</td>
		</tr>
	</table>";
	$logcnt = $i - 1;
	sqlsrv_free_stmt( $stmt );	
}else{
	echo "Connection could not be established the SQL Server.<br />";
	die( print_r( sqlsrv_errors(), true));
}

//footer 
?>	
						</form>
<br/>
</body>
</html>
			