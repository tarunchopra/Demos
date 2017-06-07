(function () {
  'use strict';

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      $(document).ready(function () {
          /*var phoneNumber = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.PhoneNumber);
          $('#debug-out').text("Phone number selected - "+ phoneNumber[0].originalPhoneString);*/
		  
		  
		  $(document).ready(function () {

            Office.context.mailbox.item.body.getAsync(

			  "text",

			  { asyncContext:"This is passed to the callback" },

			  function callback(result) {

			    if ($('#aftership-jssdk').length)
                    $('#aftership-jssdk').remove();
				
				var str = result.value; 
            	var res = str.match(/(9[1-4]{1}\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d|9[1-4]{1}\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d)|(9[1-4]{1}\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d|9[1-4]{1}\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d)/g).toString();

                $('#debug-out').prepend('<div id="as-root"></div><script>(function(e,t,n){var r,i=e.getElementsByTagName(t)[0];if(e.getElementById(n))return;r=e.createElement(t);r.id=n;r.src="//button.aftership.com/all.js";i.parentNode.insertBefore(r,i)})(document,"script","aftership-jssdk")</script>');

                $('#debug-out').append('<div class="as-track-button" style="margin-left:200px" data-tracking-number="' + res.trim() + '" data-size="small"> </div>');

				$( 'div[class^=as-container-]' ).css( {'margin-left':'100px'});

			  });

        });
      });

  };
})();