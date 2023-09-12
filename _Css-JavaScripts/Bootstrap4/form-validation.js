// Example starter JavaScript for disabling form submissions if there are invalid fields
/*
$( document ).ready(function() {
    
    // Fetch all the forms we want to apply custom Bootstrap validation styles to
    var forms = document.getElementsByClassName('needs-validation');

    // Loop over them and prevent submission
    Array.prototype.filter.call(forms, function (form) {
      form.addEventListener('submit', function (event) {
        if (form.checkValidity() === false) {
          event.preventDefault();
          event.stopPropagation();
        }

        form.classList.add('was-validated');
      }, false);
    });
    
});
*/

// https://stackoverflow.com/questions/54510043/bootstrap-and-select2-form-validation

$( document ).ready(function() {
    $(".needs-validation").on('submit', function (event) {
        $(this).addClass('was-validated');

        if ($(this)[0].checkValidity() === false) {
            event.preventDefault();
            event.stopPropagation();
            
            return false;
        } else {
            alert('form submitted');
            
            event.preventDefault();
            event.stopPropagation();
            
            return true;
        };
    });
});