const XLSX = require("xlsx");

// Crear datos de ejemplo (array de objetos)
const students = [
      { Nombre: "Ana López", Edad: 20, Curso: "Matemáticas" },
      { Nombre: "Juan Pérez", Edad: 22, Curso: "Historia" },
      { Nombre: "María Gómez", Edad: 19, Curso: "Programación" },
      { Nombre: "Luis Fernández", Edad: 21, Curso: "Física" },
      { Nombre: "Sofía Martínez", Edad: 23, Curso: "Química" },
      { Nombre: "Carlos Sánchez", Edad: 24, Curso: "Biología" },
      { Nombre: "Laura Torres", Edad: 20, Curso: "Literatura" },
      { Nombre: "Javier Ramírez", Edad: 22, Curso: "Arte" },
      { Nombre: "Patricia Díaz", Edad: 19, Curso: "Música" },
      { Nombre: "Fernando Castro", Edad: 21, Curso: "Filosofía" },
      { Nombre: "Gabriela Ruiz", Edad: 25, Curso: "Economía" },
      { Nombre: "Diego Herrera", Edad: 23, Curso: "Ingeniería" },
      { Nombre: "Valeria Ortiz", Edad: 20, Curso: "Psicología" },
      { Nombre: "Ricardo Morales", Edad: 22, Curso: "Derecho" },
      { Nombre: "Camila Vega", Edad: 19, Curso: "Medicina" },
      { Nombre: "Andrés Navarro", Edad: 21, Curso: "Arquitectura" },
      { Nombre: "Isabela Rojas", Edad: 24, Curso: "Diseño Gráfico" },
      { Nombre: "Sebastián Gil", Edad: 22, Curso: "Administración" },
      { Nombre: "Paula Medina", Edad: 23, Curso: "Marketing" },
      { Nombre: "Tomás Castillo", Edad: 20, Curso: "Ciencias Políticas" },
      { Nombre: "Lucía Vargas", Edad: 19, Curso: "Antropología" },
      { Nombre: "Mateo Paredes", Edad: 21, Curso: "Geografía" },
      { Nombre: "Daniela Fuentes", Edad: 24, Curso: "Estadística" },
      { Nombre: "Hugo Salazar", Edad: 23, Curso: "Sociología" },
      { Nombre: "Elena Cruz", Edad: 20, Curso: "Educación" },
];

// Convertir los datos en una hoja de cálculo
const worksheet = XLSX.utils.json_to_sheet(students);

// Crear un nuevo libro de trabajo (workbook)
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "Estudiantes");

// Guardar el archivo Excel
XLSX.writeFile(workbook, "lista_estudiantes.xlsx");

console.log(
  'Archivo Excel "lista_estudiantes.xlsx" creado exitosamente. Ahora puede ver la lista completa'
);



// --------------------------------------------



// Leer un archivo Excel existente
const inputFile = 'lista_estudiantes.xlsx'; // Cambia a 'data/input.xlsx' si usas otro archivo
const workbookRead = XLSX.readFile(inputFile);

// Obtener la primera hoja del archivo
const sheetName = workbookRead.SheetNames[0];
const worksheetRead = workbookRead.Sheets[sheetName];

// Convertir la hoja a un array de objetos
const data = XLSX.utils.sheet_to_json(worksheetRead);

// Mostrar los datos en la consola
console.log('Contenido del archivo Excel:');
console.log(data);