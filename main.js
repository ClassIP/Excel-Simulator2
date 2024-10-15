window.location.href = "./projects/JobReport.xlsx";
// Banco de preguntas
const bancoDePreguntas = {
    proyecto1: {
        nombre: "JobReport",
        archivo: "./projects/JobReport.xlsx",
        preguntas: [

            [
                "Freeze row 1 on the 'Last Quarter' worksheet so it remains visible while scrolling.",
                "Freeze the top 2 rows on the 'Last Quarter' worksheet to keep them visible while scrolling.",
                "On the 'Last Quarter' worksheet, freeze rows 1 through 3 so they remain visible when scrolling down the sheet.",
                "Freeze the first 4 rows on the 'Last Quarter' worksheet to maintain visibility while scrolling.",
                "Freeze rows 1 and 2 on the 'Last Quarter' worksheet to ensure they stay in view while you scroll."
            ],
            [
                "On the 'Last Quarter' worksheet, in cell A3, apply the Strikethrough font effect to the name 'Danner, Ryan.'",
                "On the 'Last Quarter' worksheet, in cell B4, apply the Superscript font effect to the company 'SynergyTech Ventures.'",
                "On the 'Last Quarter' worksheet, in cell A5, apply the Subscript font effect to the name 'Delaney, Casey.'",
                "On the 'Last Quarter' worksheet, in cell A2, apply the Strikethrough font effect to the name 'Alden, Sawyer.'",
                "On the 'Last Quarter' worksheet, in cell A3, apply the Superscript font effect to the name 'Danner, Ryan.'"
            ],
            [
                "On the 'Summary' worksheet, in the 'Trend' column, insert Line sparklines to show the trends from Year 2 through Year 5.",
                "On the 'Summary' worksheet, in the 'Trend' column, insert Column sparklines to show the trends from Year 1 through Year 5.",
                "On the 'Summary' worksheet, in the 'Trend' column, insert Win/Loss sparklines to show trends from Year 1 to Year 4.",
                "On the 'Summary' worksheet, in the 'Trend' column, insert Line sparklines to show the trends from Year 1 to Year 3.",
                "On the 'Summary' worksheet, in the 'Trend' column, insert Column sparklines for the trends between Year 2 and Year 5."
            ],
            [
                "On the 'Last Quarter' worksheet, in column F, beginning in cell F2, use a function to display each 'Job Title' from the table without retrieving duplicate entries.",
                "On the 'Last Quarter' worksheet, in column F, beginning in cell F2, use a function to display each 'Company' from the table without duplicates.",
                "On the 'Last Quarter' worksheet, in column F, beginning in cell F2, use a function to display each 'Candidate Name' without retrieving duplicate entries.",
                "On the 'Last Quarter' worksheet, in column F, beginning in cell F2, use a function to list each 'Job Title' without duplicates.",
                "On the 'Last Quarter' worksheet, in column F, starting from cell F2, use a function to extract unique 'Company' names from the table."
            ],
            [
                "On the 'Job Openings' worksheet, modify the chart to display the Primary Vertical axis title. Enter the title 'Jobs'.",
                "On the 'Job Openings' worksheet, modify the chart to display the Primary Horizontal axis title. Enter the title 'Years'.",
                "On the 'Job Openings' worksheet, modify the chart to display the Primary Vertical axis title. Enter the title 'Open Positions'.",
                "On the 'Job Openings' worksheet, modify the chart's Axes to display the Primary Horizontal Axis title. Enter 'Time Period'.",
                "On the 'Job Openings' worksheet, modify the chart's Axes to show the Primary Vertical Axis title. Enter 'Job Counts'."
            ]
        ]



    },
    // Proyecto 2 
    proyecto2: {
        nombre: "StudentsGrades",
        archivo: "./projects/StudentsGrades.xlsx",
        preguntas: [
            [
                "In the document properties, add 'Math 101' as a Company.",
                "In the document properties, add 'Algebra Basics' as a tag.",
                "In the document properties, add 'Student Work' in the Category field.",
                "In the document properties, add 'Introduction to Mathematics' as the Title.",
                "In the document properties, add 'Math Department' as the Company.",
                "In the document properties, add '2024 Semester' in the Category field.",
                "In the document properties, add 'Final Project' as the Title."
            ],
            [
                "On the 'Presentation Schedule' worksheet, modify the formula in the 'Time' column so that presentations are scheduled every 15 minutes from 8:00 AM.",
                "On the 'Presentation Schedule' worksheet, modify the formula in the 'Time' column so that presentations are scheduled every 45 minutes from 8:00 AM.",
                "On the 'Presentation Schedule' worksheet, modify the formula in the 'Time' column so that presentations are scheduled every 1 hour from 8:00 AM.",
            ],
            [
                "On the 'Grades' worksheet, in the 'Attendance' column, use conditional formatting to apply the *Green Fill with Dark Green Text* format to cells that contain values greater than 97.",
                "On the 'Grades' worksheet, in the 'Attendance' column, apply conditional formatting with *Yellow Fill with Dark Yellow Text* for cells with values greater than 90.",
                "On the 'Grades' worksheet, in the 'Attendance' column, use conditional formatting to apply *Light Red Fill with Dark Red Text* to values less than 60.",
                "On the 'Grades' worksheet, in the 'Attendance' column, apply conditional formatting with *Red Border* for values between 80 and 100.",
                "On the 'Grades' worksheet, in the 'Attendance' column, use conditional formatting to apply *Light Red Fill* to cells with values greater than 85."
            ],
            [
                "On the 'Grades' worksheet, perform a multi-level sort. Sort the table data by 'Final' (Largest to Smallest) and then by 'Student ID' (Smallest to Largest).",
                "On the 'Grades' worksheet, sort the table by 'Average Total Points' (Largest to Smallest) and then by 'Student ID' (Smallest to Largest).",
                "On the 'Grades' worksheet, sort the data first by 'Attendance' (Largest to Smallest) and then by 'Final' (Smallest to Largest).",
                "On the 'Grades' worksheet, perform a multi-level sort by 'Final' (Largest to Smallest) and then by 'Student ID' (Smallest to Largest).",
                "On the 'Grades' worksheet, sort the table by 'Final' (Smallest to Largest) and then by 'Attendance' (Largest to Smallest)."
            ],
            [
                "On the 'Grades' worksheet, in the 'Bonus' column, enter a formula that multiplies the value in the 'Attendance' column by cell H4."
            ],
            [
                "On the 'Grades' worksheet, in the 'Posted Scores' column, use a function to display the value from the 'Student ID' column, followed by the text '-Final Exam-', and the value from the 'Final' column.",
                "On the 'Grades' worksheet, in the 'Posted Scores' column, display the 'Student ID' value, followed by '-Exam Score-', and the 'Final' value.",
                "On the 'Grades' worksheet, in the 'Posted Scores' column, display the 'Student ID' value, followed by '-Grade-', and the value from the 'Final' column.",
                "On the 'Grades' worksheet, in the 'Posted Scores' column, display the 'Student ID', '-Test Result-', and the 'Final' value.",
                "On the 'Grades' worksheet, display the 'Student ID', '-Final-', 'Final' score, followed by the 'First Name'."
            ],
            [
                "On the 'Attendance Analysis' worksheet, add the alt text description 'Attendance chart' to the chart.",
                "On the 'Attendance Analysis' worksheet, add the alt text description 'Student Attendance Overview' to the chart.",
                "On the 'Attendance Analysis' worksheet, add the alt text description 'Attendance Performance' to the chart.",
                "On the 'Attendance Analysis' worksheet, add the alt text description 'Weekly Attendance Summary' to the chart.",
                "On the 'Attendance Analysis' worksheet, add the alt text description 'Class Attendance Analysis' to the chart."
            ]



        ]
    },

    // Proyecto 3 
    proyecto3: {
        nombre: "BookPublishing",
        archivo: "./projects/BookPublishing.xlsx",
        preguntas: [
            // Importing Data and Formatting as a Table
            [
                "On the 'Out of Print' worksheet, starting at cell A3, import data from the 'OutOfPrint' text file located in the Documents folder. Ensure the table does not use the first row of the data source as headers, but as data.",
                "On the 'Out of Print' worksheet, starting from cell A3, import data from the 'OutOfPrint' text file in the Documents folder, and format the first row as data, not headers.",
                "In the 'Out of Print' worksheet, import data from the 'OutOfPrint' text file starting at cell A4. Use the first row as data, not headers.",
                "On the 'Out of Print' sheet, import data from the 'OutOfPrint' text file located in Documents, starting at cell B3. Ensure the first row is included as data.",
                "On the 'Out of Print' sheet, starting at cell A3, import data from the 'OutOfStock' text file in Documents, making sure the first row is part of the data, not headers."
            ],

            // Setting Horizontal Text Alignment Variations
            [
                "On the 'Inventory' worksheet, set the horizontal text alignment of cell I3 to Center.",
                "In the 'Inventory' worksheet, set the horizontal text alignment of cell J3 to Center.",
                "On the 'Inventory' sheet, set the horizontal text alignment for cells I3 and J3 to Center.",
                "In the 'Inventory' worksheet, change the horizontal alignment of cell I3 to Center Across Selection.",
                "On the 'Inventory' sheet, apply Center Across Selection to the text in cell J3."
            ],

            // Adding a Column to the Table with Header Change
            [
                "On the 'Inventory' worksheet, add only column G to the 'Year-End Inventory' table and rename the header to 'Total Value'.",
                "On the 'Inventory' worksheet, insert only column G into the 'Year-End Inventory' table and change the header to 'Accumulated Value'.",
                "In the 'Inventory' sheet, add column G to the 'Year-End Inventory' table and rename the header to 'Net Value'.",
                "On the 'Inventory' worksheet, add column G to the 'Year-End Inventory' table and update the header to 'Total Revenue'.",
                "In the 'Inventory' sheet, insert column G into the 'Year-End Inventory' table and rename the header to 'Final Value'."
            ],

            // Finding Maximum/Minimum Value Using a Function
            [
                "On the 'Inventory' worksheet, in cell F26, use a function to display the minimum value from the 'Unit Price' column of the 'Year-End Inventory' table.",
                "In the 'Inventory' worksheet, use a function in cell F26 to calculate the lowest value from the 'Unit Price' column.",
                "On the 'Inventory' sheet, in cell F26, use a function to show the minimum value in the 'Unit Price' column of the 'Year-End Inventory' table.",
                "In the 'Inventory' worksheet, in cell F26, apply a function to return the smallest value in the 'Unit Price' column of the 'Year-End Inventory' table.",
                "On the 'Inventory' sheet, use a function in cell F26 to find the lowest value in the 'Unit Price' column."
            ],

            // Creating a 3D Chart with Variations in Chart Type
            [
                "On the 'Comparison' worksheet, using the data from 'Price Comparison by Regions', create a 3D Clustered Bar  Chart that shows the 'Total Value' data for each 'Region'. Display the regions as a legend and add the title 'Total Value'.",
                "On the 'Comparison' worksheet, using the data from 'Price Comparison by Regions', generate a 3D Column Chart for the 'Total Value' of each region, displaying the regions in the legend and the title 'Total Value'.",
                "In the 'Comparison' worksheet, create a 3D Pie Chart using the 'Total Value' data from 'Price Comparison by Regions'. Show the regions as a legend and add the title 'Total Value'.",
                "On the 'Comparison' worksheet, generate a 3D Area Chart displaying the 'Total Value' data for each 'Region' from 'Price Comparison by Regions'. Show the regions in the legend with the title 'Total Value'.",
                "In the 'Comparison' worksheet, create a 3D Doughnut Chart using the 'Total Value' data for each region from 'Price Comparison by Regions', with a legend for the regions and the title 'Total Value'."
            ],

            // Modifying a Chart with Data Label Variations
            [
                "On the 'Inventory' worksheet, modify the chart to display the series values as data labels inside base of each column.",
                "On the 'Inventory' sheet, adjust the chart to display the series values as data labels Outside the end of each column.",
                "In the 'Inventory' worksheet, change the chart to display the series values as data labels inside the end of each column.",
                "On the 'Inventory' sheet, modify the chart to show data labels inside the columns instead of at the end.",
                "In the 'Inventory' worksheet, update the chart to display the series values as data labels at the center of each column."
            ]
        ]
    },

    // Proyecto 4 
    proyecto4: {
        nombre: "Proximamente",
        archivo: "",
        preguntas: [

        ]
    },

    // Proyecto 5
    proyecto5: {
        nombre: "Proximamente",
        archivo: "",
        preguntas: [

        ]
    },

    // Proyecto 6
    proyecto6: {
        nombre: "Proximamente",
        archivo: "",
        preguntas: [

        ]
    },
};

// Control del proyecto y preguntas seleccionadas
const totalProjects = Object.keys(bancoDePreguntas).length;
let currentProjectIndex = 0;
let currentQuestionIndex = 0;
let selectedVariants = [];
let timer;
let secondsRemaining = 50 * 60;

// Función para seleccionar y guardar una variante aleatoria por pregunta
function selectRandomVariants(project) {
    const projectData = bancoDePreguntas[project];
    selectedVariants = projectData.preguntas.map((variants) => {
        const randomVariant = variants[Math.floor(Math.random() * variants.length)];
        return randomVariant;
    });
}

// Función para cargar preguntas y actualizar la interfaz
function loadProjectQuestions() {
    const projectKey = Object.keys(bancoDePreguntas)[currentProjectIndex];
    const projectData = bancoDePreguntas[projectKey];
    const totalQuestions = projectData.preguntas.length;

    // Actualizar título del proyecto
    document.getElementById("project-title").textContent = `Project ${currentProjectIndex + 1} of ${totalProjects}: ${projectData.nombre}`;

    // Actualizar barra de navegación de preguntas
    const navigationBar = document.getElementById("navigation-bar");
    navigationBar.innerHTML = ''; // Limpiar botones antiguos

    const prevBtn = document.createElement('button');
    prevBtn.textContent = "◄";
    prevBtn.addEventListener("click", goPrevQuestion);
    navigationBar.appendChild(prevBtn);

    for (let i = 0; i < totalQuestions; i++) {
        const btn = document.createElement('button');
        btn.classList.add("question-btn");
        btn.textContent = (i + 1);
        btn.addEventListener("click", () => {
            currentQuestionIndex = i;
            loadQuestion(currentQuestionIndex);
        });
        if (i === 0) btn.classList.add("active");
        navigationBar.appendChild(btn);
    }

    const nextBtn = document.createElement('button');
    nextBtn.textContent = "►";
    nextBtn.addEventListener("click", goNextQuestion);
    navigationBar.appendChild(nextBtn);

    // Cargar la primera pregunta
    loadQuestion(currentQuestionIndex);
}

// Función para cargar la pregunta almacenada
function loadQuestion(index) {
    const projectKey = Object.keys(bancoDePreguntas)[currentProjectIndex];
    const questionText = selectedVariants[index];
    document.getElementById("question-text").textContent = questionText;

    // Actualizar la pregunta activa en la barra de navegación
    const buttons = document.querySelectorAll(".question-btn");
    buttons.forEach((btn, i) => {
        btn.classList.toggle("active", i === index);
    });
}

// Navegar entre preguntas
function goPrevQuestion() {
    if (currentQuestionIndex > 0) {
        currentQuestionIndex--;
        loadQuestion(currentQuestionIndex);
    }
}

function goNextQuestion() {
    const projectKey = Object.keys(bancoDePreguntas)[currentProjectIndex];
    if (currentQuestionIndex < bancoDePreguntas[projectKey].preguntas.length - 1) {
        currentQuestionIndex++;
        loadQuestion(currentQuestionIndex);
    }
}

// Función para iniciar el temporizador regresivo
function startTimer() {
    timer = setInterval(() => {
        secondsRemaining--;
        displayTime();

        if (secondsRemaining <= 0) {
            clearInterval(timer);
            alert("Time is up! The project will be submitted.");
            submitProject();
        }
    }, 1000);
}

// Función para mostrar el tiempo en la pantalla
function displayTime() {
    const timerElement = document.getElementById("timer");
    timerElement.textContent = formatTime(secondsRemaining);
}

// Función para formatear el tiempo en mm:ss
function formatTime(seconds) {
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = seconds % 60;
    return `${minutes.toString().padStart(2, '0')}:${remainingSeconds.toString().padStart(2, '0')}`;
}

// Función para enviar el proyecto y cambiar al siguiente
function submitProject() {

    // Obtener la ruta del archivo actual y redirigir
    const projectKey = Object.keys(bancoDePreguntas)[currentProjectIndex + 1];
    const archivoProyecto = bancoDePreguntas[projectKey].archivo;
    window.location.href = archivoProyecto; // Redirigir al archivo del proyecto actual

    // Cambiar al siguiente proyecto si hay más
    if (currentProjectIndex < totalProjects - 1) {
        currentProjectIndex++;
    } else {
        currentProjectIndex = 0; // Volver al primer proyecto si se termina la lista
    }

    // Reiniciar el estado para el siguiente proyecto
    currentQuestionIndex = 0;
    selectRandomVariants(projectKey); // Seleccionar nuevas variantes
    loadProjectQuestions(); // Cargar las preguntas del nuevo proyecto
}

// Función para detener el temporizador
function stopTimer() {
    clearInterval(timer);
}

// Inicializar con el primer proyecto
selectRandomVariants(Object.keys(bancoDePreguntas)[currentProjectIndex]);
loadProjectQuestions();
startTimer(); // Iniciar el temporizador cuando el proyecto cargue

// Enviar el proyecto manualmente al hacer clic en el botón
document.getElementById("submit-project").addEventListener("click", () => {
    alert("Proyecto enviado.");
    submitProject();
});