chrome.extension.sendMessage({}, function(response) {
	var readyStateCheckInterval = setInterval(function() {
	if (document.readyState === "complete") {
		clearInterval(readyStateCheckInterval);

		// ----------------------------------------------------------
		// This part of the script triggers when page is done loading
		// console.log("Hello. This message was sent from scripts/inject.js");
		// ----------------------------------------------------------

		// TODO: check if we're on a board

		var ttt = {

			init: function() {
				// $('body').on('click', '.js-ttt-card', function() {
				// 	alert("CLICK TODO");
				// 	$('.card-detail-window .comment-box textarea').html("@h 8:00 Example acvitity log");
				// });
			},
			
			loadHours:  function(mutationType) {
				var $detailWindow = $('.card-detail-window');
				if (mutationType == 'childList' && $detailWindow.size() > 0
						&& $detailWindow.find('#ttt-actions').size() == 0) {
						
					// Add time tracking button
					$detailWindow.find('.other-actions h3').after(ttt.renderActions());
					$('#ttt-actions .js-ttt-card').on('click', function() { ttt.toggle(); });

					// Listen for comment updates
					// TODO: also listen for new comments
					commentsObserver = new MutationObserver(function(mutations) {
						mutations.forEach(function(mutation) {
							if (mutation.type == 'childList') {
								// Update hours log
								$('#ttt-actions .js-ttt-card .sum').text(ttt.getHoursSum());
							}
						});
					});
					commentsObserver.observe(document.querySelector('.js-list-actions'), { 
						attributes: false, childList: true, characterData: false 
					});
				}
				else if ($detailWindow.size() == 0 && commentsObserver) {
					// Stop listening for comment updates
					commentsObserver.disconnect();
					commentsObserver = null;
				}
			},

			toggleState: true,

			toggle: function() {
				if (ttt.toggleState) {
					$('.ttt-comment').removeClass('ttt-comment-hidden');
				}	
				else {
					$('.ttt-comment').addClass('ttt-comment-hidden');
				}
				ttt.toggleState = !ttt.toggleState;
			},

			renderActions: function() {
				ttt.toggleState = true;

				return '<div id="ttt-actions">'
						+'<a class="button-link js-ttt-card" href="#" title="@h 8:00 Example acvitity log">'
						+'<span class="icon-sm icon-clock"></span> Hours <span class="sum">'+ttt.getHoursSum()+"</span>"
						+'</a>'
						+'</div>';
			},

			getHoursSum: function() {
				// Sum (minutes)
				var sum = 0;
				$('.window-wrapper .current-comment').each(function() {
					var $this = $(this);
					var match = $this.text().match(/@h ([0-9]{1,2}):([0-9]{1,2}) (.*)/);
					if (match) {
						sum += parseInt(match[1])*60 + parseInt(match[2]);

						$this.parents('.mod-comment-type')
								.addClass('ttt-comment')
								.addClass('ttt-comment-hidden');
					}  
				});

				// Format
				var sec_num = sum * 60;
				var hours   = Math.floor(sec_num / 3600);
				var minutes = Math.floor((sec_num - (hours * 3600)) / 60);
				var seconds = sec_num - (hours * 3600) - (minutes * 60);
				if (hours   < 10) {hours   = "0"+hours;}
				if (minutes < 10) {minutes = "0"+minutes;}
				if (seconds < 10) {seconds = "0"+seconds;}
				
				return hours+':'+minutes; //+':'+seconds;
			}
		};

		// Replaced by 
//		$(document).on('ready', function() {
			ttt.init();
//		});

		var commentsObserver = null;
		var cardDetailsObserver = new MutationObserver(function(mutations) {
			mutations.forEach(function(mutation) {
				ttt.loadHours(mutation.type);
			});
		});
		cardDetailsObserver.observe(document.querySelector('.window-wrapper'), { 
			attributes: false, childList: true, characterData: false 
		});
		// cardDetailsObserver.disconnect();
		
		if (window.location.pathname.split(/\//)[1] /*b=board, c=card*/ == 'c') {
			ttt.loadHours('childList');
		}
	}
	}, 10);
});
