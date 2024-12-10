    // Initialization when Office JS and JQuery are ready
    Office.onReady(() => {
        $((info) => {
            $('#generate-image').on('click', insertGeneratedImage);
        });
    });

    function setSlideSize(width, height) {
       Office.context.document.getActiveViewAsync((result) => {
           if (result.status === Office.AsyncResultStatus.Succeeded) {
               const view = result.value;
               if (view === Office.ActiveView.ViewType.Slide) {
                   Office.context.document.setSelectedDataAsync(
                       {
                           width: width,
                           height: height
                       },
                       { coercionType: Office.CoercionType.SlideSize },
                       (asyncResult) => {
                           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                               showNotification('Error in setSlideSize:', '"' + asyncResult.error.message + '"');
                           }
                       }
                   );
               }
           } else {
               showNotification('Error in setSlideSize:', '"' + result.error.message + '"');
           }
       });
    }

    function insertGeneratedImage() {
       const imagePicker = document.getElementById("imagePicker");
       const file = imagePicker.files[0];

       if (file && (file.type === "image/png" || file.type === "image/jpeg" || file.type === "image/svg+xml")) {
           const reader = new FileReader();
           reader.onload = function(event) {
               const base64Image = event.target.result.split(",")[1];
               Office.context.document.setSelectedDataAsync(base64Image, { coercionType: Office.CoercionType.Image }, function (asyncResult) {
                   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      showNotification('Error in insertImage:', '"' + asyncResult.error.message + '"');
                   }
               });
           };
           reader.readAsDataURL(file);
       } else {
           showNotification('Error in insertImage:', "Please select a valid PNG, JPEG, or SVG image.");
       }
   }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        //$("#notificationHeader").text(header);
        //$("#notificationBody").text(content);
        //messageBanner.showBanner();
        //messageBanner.toggleExpansion();
    }
