var MultiRefinement = MultiRefinement || {};

MultiRefinement.SubmitRefinement = function (name, control) {
    // Get the Refiner Control Element from the page
    var refinerElm = document.getElementById(name + '-MultiRefiner');
    if (refinerElm) {
        // Retrieve all the checkboxes from the control
        var checkboxElms = refinerElm.getElementsByTagName('input');
        // Retrieve the operator dropdown element
        var operatorElm = document.getElementsByName(name + '-Operator');
        // Get the operator value
        var operator = 'OR';
        for (var i=0; i < operatorElm.length; i++) {
            var elm = operatorElm[i]
            if (elm.checked) {
                operator = elm.value;
            }
        }
        
        // Create a new array
        var refiners = [];
        // Loop over each checkbox
        for (var i = 0; i < checkboxElms.length; i++) {
            var elm = checkboxElms[i];
            // Check if the checkbox is checked
            if (elm.checked) {
                // Append the refiner value to the array
                Srch.U.appendArray(refiners, elm.value);
            }
        };

        // Call the refinement method with the array of refiners
        control.updateRefinementFilters(name, refiners, operator, false, null);
    }
};