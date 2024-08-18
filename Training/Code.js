// Global column map
var wlColMap = {
    EXERCISE:        [ 1, 'A'],   
    SETS:            [ 2, 'B'],
    REPS:            [ 3, 'C'],
    WEIGHT:          [ 4, 'D'],
    UNIT:            [ 5, 'E'],
    BODY:            [ 6, 'F'],
    TOTAL:           [ 7, 'G'],
    INTENSITY_SCORE: [ 8, 'H'],
    DATE:            [ 9, 'I'],
    NOTES:           [10, 'J']
};                     

var BodyAreas = {
    Arms: "Arms",
    Chest: "Chest",
    Upper: "Upper Body",
    Legs: "Legs",
    Back: "Back",
};

/*
    name: Name of the exercise 
    bodyArea: Area the exercise is targeting
    weightPercentage: What percentage of the weight set is actually being used
    maxWeight: Maximum target weight
    exerciseFactor: How difficult the exercise is - from .1 - 2
*/

var Exercises = [ 
    {
        name: "Chest Fly Machine",
        bodyArea: BodyAreas.Chest,
        weightPercentage: .5,
        maxWeight: 150,
        exerciseFactor: 1 
    },

    {
        name: "Bicep Curl Machine",
        bodyArea: BodyAreas.Arms,
        weightPercentage: .5,
        maxWeight: 80,
        exerciseFactor: 1 
    },

    {
        name: "Bicep Curl Free Weights",
        bodyArea: BodyAreas.Arms,
        weightPercentage: 1,
        maxWeight: 50,
        exerciseFactor: 1 
    },

    {
        name: "Shoulder Press Free Weights",
        bodyArea: BodyAreas.Upper,
        weightPercentage: 1,
        maxWeight: 70,
        exerciseFactor: 1 
    },

    {
        name: "Push Up",
        bodyArea: BodyAreas.Upper,
        weightPercentage: .65,
        maxReps: 100,
        exerciseFactor: 1.5 
    },

    {
        name: "Chin Up",
        bodyArea: BodyAreas.Upper,
        weightPercentage: .95,
        maxReps: 20,
        exerciseFactor: 1.5 
    },

    {
        name: "Pull Up",
        bodyArea: BodyAreas.Upper,
        weightPercentage: .95,
        maxReps: 15,
        exerciseFactor: 1.7 
    },

    {
        name: "Calf Raise",
        bodyArea: BodyAreas.Legs,
        weightPercentage: .6,
        maxReps: 200,
        exerciseFactor: .7 
    },

    {
        name: "Bicep Pull Down",
        bodyArea: BodyAreas.Arms,
        weightPercentage: 1,
        maxWeight: 150,
        exerciseFactor: 1 
    },

    {
        name: "Hammer Curl Free Weights",
        bodyArea: BodyAreas.Arms,
        weightPercentage: 1,
        maxWeight: 50,
        exerciseFactor: 1 
    },

    {
        name: "Incline Shoulder Press",
        bodyArea: BodyAreas.Upper,
        weightPercentage: 1,
        maxWeight: 50,
        exerciseFactor: 1 
    },

    {
        name: "Lat Pull Down",
        bodyArea: BodyAreas.Back,
        weightPercentage: 1,
        maxWeight: 200,
        exerciseFactor: 1 
    },

    {
        name: "Leg Curl Machine",
        bodyArea: BodyAreas.Legs,
        weightPercentage: .5,
        maxWeight: 200,
        exerciseFactor: 1 
    },

    {
        name: "Leg Extension Machine",
        bodyArea: BodyAreas.Legs,
        weightPercentage: .5,
        maxWeight: 200,
        exerciseFactor: 1 
    },

    {
        name: "Pectoral Fly Machine",
        bodyArea: BodyAreas.Chest,
        weightPercentage: .5,
        maxWeight: 200,
        exerciseFactor: 1 
    },

    {
        name: "Shoulder Press Machine",
        bodyArea: BodyAreas.Upper,
        weightPercentage: .5,
        maxWeight: 200,
        exerciseFactor: 1 
    },

    {
        name: "Tricep Extension Machine",
        bodyArea: BodyAreas.Arms,
        weightPercentage: .5,
        maxWeight: 150,
        exerciseFactor: 1 
    },

    {
        name: "Back Row Machine",
        bodyArea: BodyAreas.Back,
        weightPercentage: .5,
        maxWeight: 150,
        exerciseFactor: 1 
    },

];

var scalingFactor = .5;

function setExerciseDropdown(sheet) {

    // Define the range (column A in this case, adjust as necessary)
    var range = sheet.getRange(wlColMap.EXERCISE[1] + "2:" + wlColMap.EXERCISE[1] + "10000");

    const exercises = Exercises.map(exercise => exercise.name);

    // Create the data validation rule
    var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(exercises, true)
        .setAllowInvalid(true)
        .build();

    // Apply the data validation rule to the range
    range.setDataValidation(rule);
}

function setBodyDropdown(sheet) {

    // Define the range (column A in this case, adjust as necessary)
    var range = sheet.getRange(wlColMap.BODY[1] + "2:" + wlColMap.BODY[1] + "10000");

    const bodyAreasArray = Object.values(BodyAreas);

    // Create the data validation rule
    var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(bodyAreasArray, true)
        .setAllowInvalid(true)
        .build();

    // Apply the data validation rule to the range
    range.setDataValidation(rule);
}

function setUnitDropdown(sheet) {

    // Define the range (column A in this case, adjust as necessary)
    var range = sheet.getRange(wlColMap.UNIT[1] + "2:" + wlColMap.UNIT[1] + "10000");

    // Create the data validation rule
    var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["lbs", "kg"], true)
        .setAllowInvalid(true)
        .build();

    // Apply the data validation rule to the range
    range.setDataValidation(rule);
}

function onOpen(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    setExerciseDropdown(sheet);
    setBodyDropdown(sheet);
    setUnitDropdown(sheet);

    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')  // The name of your custom menu
      .addItem('Process All Rows', 'processAllRows')  // Menu item label and function name
      .addToUi();
}

function processAllRows() {

    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getActiveSheet();
    var numRows = sheet.getLastRow();

    // Loop through each row
    for (var i = 2; i <= numRows ; i++) {

        // Create a mock event object for the onEdit trigger
        var e = {
            range: sheet.getRange(i, 1),  // Mock the range as the first column for example
            value: sheet.getRange(i, 1).getValue(),            // Mock value from the first column of each row
            source: spreadSheet,
            user: Session.getActiveUser()
        };

        // Call onEdit function with the mock event object
        onEdit(e);
    }
}

function onEdit(e) {
    var range = e.range;

    // Check if the edited cell is in the UNIT column
    if (range.getColumn() == wlColMap.UNIT[0]) {
        weightConvert(e);
    }

    // Check if the edited cell is in the Exercise
    if (range.getColumn() == wlColMap.EXERCISE[0]) {
        exerciseConverter(e);
    }

    if (range.getColumn() == wlColMap.WEIGHT[0]) {
        weightUnitDefault(e);
    }

    setDate(e);
    calcTotal(e);
}

function weightConvert(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var value = range.getValue();

    var weightCell = sheet.getRange(range.getRow(), wlColMap.WEIGHT[0]); // Get the corresponding cell in Weight column 
    var weight = weightCell.getValue();

    if (value.toLowerCase() == "kg") {
        var convertedWeight = Math.trunc(weight * 2.20462); // Convert kg to lbs
        weightCell.setValue(convertedWeight);
        range.setValue("lbs"); // Update unit to lbs
    }
}

function weightUnitDefault(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    var unitCell = sheet.getRange(range.getRow(), wlColMap.UNIT[0]); // Get the corresponding cell in Weight column 
    unitCell.setValue("lbs");
}

function exerciseConverter(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var value = range.getValue();

    exercise = Exercises.find(obj => obj.name === value);

    var bodyCell = sheet.getRange(range.getRow(), wlColMap.BODY[0]);  
    bodyCell.setValue(exercise.bodyArea);
}

function calcTotal(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    if(range.getRow() === 1) { return; }

    var totalCell = sheet.getRange(range.getRow(), wlColMap.TOTAL[0]);  
    totalCell.setValue("");

    var weightCell = sheet.getRange(range.getRow(), wlColMap.WEIGHT[0]);  
    var weight = weightCell.getValue();

    if (weight !== "") {
        var setCell = sheet.getRange(range.getRow(), wlColMap.SETS[0]);  
        var set = setCell.getValue();

        if (set !== "") {
            var repCell = sheet.getRange(range.getRow(), wlColMap.REPS[0]);  
            var rep = repCell.getValue();

            if (rep !== "") {
                var exerciseCell = sheet.getRange(range.getRow(), wlColMap.EXERCISE[0]);  
                var exercise = Exercises.find(obj => obj.name === exerciseCell.getValue());

                totalCell.setValue(Math.trunc(weight * set * rep * exercise.weightPercentage));
                calcIntensity(e, exercise, weight, set, rep);
            }
        }
    }
}

function setDate(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    var value = range.getValue();
    var col = range.getColumn();

    if(col === wlColMap.DATE[0] || sheet.getName() !== "Weightlifting") { return; }

    var dateCell = sheet.getRange(range.getRow(), wlColMap.DATE[0]);  

    if(dateCell.getValue() !== "") {return;}

    var today = new Date();
    dateCell.setValue(today);
}

function calcIntensity(e, exercise, weight, set, rep) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    var intensityCell = sheet.getRange(range.getRow(), wlColMap.INTENSITY_SCORE[0]);  

    if(exercise.maxWeight > 0) {
        var intensity = ((weight * exercise.weightPercentage * set * rep) / exercise.maxWeight) * exercise.exerciseFactor * scalingFactor;
        intensityCell.setValue(intensity);
    }

}
