function func()
{
	try{
		console.log('func started');
		try{
			var myWebSocket = new WebSocket("ws://127.0.0.1:5001");
		}catch(err){
			setTimeout(function() {
				location.reload();
			}, 3000);
		}
		myWebSocket.onerror=function(event){	
			setTimeout(function() {
				location.reload();
				}, 5000);
		}
		myWebSocket.onclose=function(event){	
			setTimeout(function() {
				location.reload();
				}, 5000);
		}
		if(window.location.href.substr(-12)=='default.aspx')
		{
			myWebSocket.onmessage = function(evt) { 
				try{
					console.log('received:'+evt.data)
					if(evt.data != 'NA' && evt.data != 'DONE' && evt.data != '')
					{
						document.getElementsByName('ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$_TextBoxParcelNumber')[0].click()
						document.getElementsByName('ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$_TextBoxParcelNumber')[0].value=evt.data;
						
						setTimeout(function() {
							document.getElementById('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__ButtonSearch').click();
						}, 3000);
						
						setTimeout(function() {
							document.getElementById('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__GridViewResults_ctl02__LinkButtonDetails').click();
						}, 7000);
						setTimeout(function() {
							document.getElementById('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__GridViewResults_ctl02__LinkButtonDetails').click();
						}, 10000);
						setTimeout(function() {
							document.getElementById('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__GridViewResults_ctl02__LinkButtonDetails').click();
						}, 13000);
						setTimeout(function() {
							if(document.getElementById('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__GridViewResults_ctl02__LinkButtonDetails')==null){
								try{
									myWebSocket.send((evt.data+'|NA'));
								}catch(err){
									location.reload();
									}
							}
						}, 15000);
						setTimeout(function() {
							try{
									myWebSocket.send("PARCEL");
								}catch(err){
										location.reload();
										}
						}, 18000);
					}else{
						setTimeout(function() {
							try{
									myWebSocket.send("PARCEL");
								}catch(err){
										location.reload();
										}
						}, 2000);
					}
				}catch(err){
					location.reload();
				}
			};
			
			setTimeout(function() {
				try{
					myWebSocket.send("PARCEL");
				}catch(err){
						location.reload();
						}
			}, 2000);
		}
		else 
		{
			myWebSocket.onmessage = function(evt) { 
			if(evt.data == 'DONE')
				{
					window.location.href="../search/default.aspx";
				}
			};

			parcel_id = (window.location.href.substr(-6));
			parcel = document.getElementById('ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__Header').textContent.replace('Parcel # ','');
			
			setTimeout(function() {
				myWebSocket.send((parcel+'|'+parcel_id));
			}, 1000);
			setInterval(function(){
				location.reload();
			}, 5000);
		}
	}catch(err){
		setTimeout(function() {
			location.reload();
		}, 3000);
	}
}

func();

setTimeout(function() {
			location.reload();
}, 60000);
