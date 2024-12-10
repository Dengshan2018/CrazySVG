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

    function blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => {
                resolve(reader.result.split(',')[1]);
            };
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }

    function insertGeneratedImage() {
       serverUrl=$('#serverurl').val();
       method=$('#method').val();
       if (method == "POST") {
           params = {prompt: $('#prompt').val()};
           opts = {
                method: method,
                headers: {
                  'Access-Control-Allow-Origin':'*',
                  'Access-Control-Allow-Methods':'POST,PATCH,OPTIONS'
                },
                body: JSON.stringify(params),
              };
       }
       else {
           opts = {
                method: method,
                headers: {
                  'Access-Control-Allow-Origin':'*',
                  'Access-Control-Allow-Methods':'POST,PATCH,OPTIONS'
                }
       }
       
       fetch(serverUrl, opts)
          .then(response => response.blob())
          .then(blob => blobToBase64(blob).then( base64Image => {
            Office.context.document.setSelectedDataAsync(base64Image, { coercionType: Office.CoercionType.Image }, function (asyncResult) {
                   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      showNotification('Error in insertImage:', '"' + asyncResult.error.message + '"');
                   }
               });
          }))
          .catch(error => {
            showNotification('File download failed:', error);
          });
   }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        console.log(header + content);
    }
