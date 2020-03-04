const app = function () {
	const API_BASE = 'https://script.google.com/macros/s/AKfycbyP5Rifn7Q05Qcd7CTfm-AOouFHHvUAvCVVuKSfQu-LCqJocP8/exec';
	const API_KEY = 'abcdef';
	const CATEGORIES = ['general', 'financial', 'technology', 'marketing'];

	const state = {activePage: 1, activeCategory: null};
	const page = {};

var beurl = "https://docs.google.com/spreadsheets/d/1RV28QjrRuOhQ1NT6ivLPljzy_yu8o-SBrxixYUnH_Pg/edit#gid=141623591";


function loadEnquiry2(){

		var ss = SpreadsheetApp.openByUrl(url);
		var bess = SpreadsheetApp.openByUrl(beurl);
		var ws = bess.getSheetByName("Option");

		var list = ws.getRange(1, 1,ws.getRange("A1").getDataRegion().getLastRow(),1).getValues();
		var list2 = ws.getRange(1, 3,ws.getRange("C1").getDataRegion().getLastRow(),1).getValues();
		var htmlListArray = list.map(function(r){ return '<option>' + r[0] + '</option>'; }).join('');
		var htmlListArray2 = list2.map(function(t){ return '<option>' + t[0] + '</option>'; }).join('');
		Logger.log(ws.getRange("A1").getDataRegion().getLastRow());
		Logger.log(ws.getRange("C1").getDataRegion().getLastRow());


	//  Table Content - All Txn
	   var wsData = bess.getSheetByName("CalculatedData");
	   var Tlist = wsData.getRange(1, 1,wsData.getRange("A1").getDataRegion().getLastRow(),6).getValues();
	  Logger.log("Tlist: " + Tlist);
	  var rNum = wsData.getRange("A1").getDataRegion().getLastRow();
	  var TArray = [];

	  for (let i=1; i<rNum; i++){
	    if(Tlist[i][4] == ""){

	    } else if ( Tlist[i][1] == "新增倉存" || Tlist[i][1] == "減少倉存" ) {
	      var TListArray = '<tr>' + '<td>' + (Tlist[i][0]).toLocaleDateString() +
	          '</td>' + '<td>' + Tlist[i][1] +
	            '</td>' + '<td>' + Tlist[i][2] +
	              '</td>' + '<td>' + Tlist[i][3] +
	                '</td>' + '<td>' + Tlist[i][4] +
	                  '</td>'  + '<td>' + (Tlist[i][5]).toLocaleDateString() + '</td>' + '</tr>';
	      TArray.push(TListArray);
	    }
	  }

	// Table Content - Summary
	  var wsData2 = bess.getSheetByName("CalculatedDataPTable");
	  var Tlist2 = wsData2.getRange(1, 1,wsData2.getRange("A1").getDataRegion().getLastRow(),3).getValues();
	  Logger.log("Tlist2: " + Tlist2);
	  var rNum2 = wsData2.getRange("A1").getDataRegion().getLastRow();
	  var TArray2 = [];

	  for (let i=1; i<rNum2; i++){
	    if(Tlist2[i][0] == "" || Tlist2[i][0] == "地點" ){

	    } else {
	      var TListArray2 = '<tr>' + '<td>' + (Tlist2[i][0]) +
	          '</td>' + '<td>' + Tlist2[i][1] +
	            '</td>' + '<td>' + Tlist2[i][2] + '</td>' + '</tr>';
	      TArray2.push(TListArray2);
	    }
	  }


	  var TArr = TArray.join('');
	  Logger.log(TArr);
	  var TArr2 = TArray2.join('');
	  Logger.log(TArr);

	  var tmp = HtmlService.createTemplateFromFile('enquiry');
	  tmp.list = htmlListArray;
	  tmp.list2 = htmlListArray2;
	  tmp.Tlist = TArr;
	  tmp.Tlist2 = TArr2;

console.log(tmp);

	  return tmp.evaluate()
	  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

	function init () {
		page.notice = document.getElementById('notice');
		page.filter = document.getElementById('filter');
		page.container = document.getElementById('container');

		_buildFilter();
		_getNewPosts();
	}

	function _getNewPosts () {
		page.container.innerHTML = '';
		_getPosts();
	}

	function _getPosts () {
		_setNotice('Loading posts');

		fetch(_buildApiUrl(state.activePage, state.activeCategory))
			.then((response) => response.json())
			.then((json) => {
				if (json.status !== 'success') {
					_setNotice(json.message);
				}

				_renderPosts(json.data);
				_renderPostsPagination(json.pages);
			})
			.catch((error) => {
				_setNotice('Unexpected error loading posts');
			})
	}

	function _buildFilter () {
	    page.filter.appendChild(_buildFilterLink('no filter', true));

	    CATEGORIES.forEach(function (category) {
	    	page.filter.appendChild(_buildFilterLink(category, false));
	    });
	}

	function _buildFilterLink (label, isSelected) {
		const link = document.createElement('button');
	  	link.innerHTML = _capitalize(label);
	  	link.classList = isSelected ? 'selected' : '';
	  	link.onclick = function (event) {
	  		let category = label === 'no filter' ? null : label.toLowerCase();

			_resetActivePage();
	  		_setActiveCategory(category);
	  		_getNewPosts();
	  	};

	  	return link;
	}

	function _buildApiUrl (page, category) {
		let url = API_BASE;
		url += '?key=' + API_KEY;
		url += '&page=' + page;
		url += category !== null ? '&category=' + category : '';

		return url;
	}

	function _setNotice (label) {
		page.notice.innerHTML = label;
	}

	function _renderPosts (posts) {
		posts.forEach(function (post) {
			const article = document.createElement('article');
			article.innerHTML = `
				<h2>${post.title}</h2>
				<div class="article-details">
					<div>By ${post.author} on ${_formatDate(post.timestamp)}</div>
					<div>Posted in ${post.category}</div>
				</div>
				${_formatContent(post.content)}
			`;
			page.container.appendChild(article);
		});
	}

	function _renderPostsPagination (pages) {
		if (pages.next) {
			const link = document.createElement('button');
			link.innerHTML = 'Load more posts';
			link.onclick = function (event) {
				_incrementActivePage();
				_getPosts();
			};

			page.notice.innerHTML = '';
			page.notice.appendChild(link);
		} else {
			_setNotice('No more posts to display');
		}
	}

	function _formatDate (string) {
		return new Date(string).toLocaleDateString('en-GB');
	}

	function _formatContent (string) {
		return string.split('\n')
			.filter((str) => str !== '')
			.map((str) => `<p>${str}</p>`)
			.join('');
	}

	function _capitalize (label) {
		return label.slice(0, 1).toUpperCase() + label.slice(1).toLowerCase();
	}

	function _resetActivePage () {
		state.activePage = 1;
	}

	function _incrementActivePage () {
		state.activePage += 1;
	}

	function _setActiveCategory (category) {
		state.activeCategory = category;

		const label = category === null ? 'no filter' : category;
		Array.from(page.filter.children).forEach(function (element) {
  			element.classList = label === element.innerHTML.toLowerCase() ? 'selected' : '';
  		});
	}

	return {
		init: init
 	};
}();
