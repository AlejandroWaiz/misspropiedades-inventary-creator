const {
    Document,
    Header,
    Footer,
    Paragraph,
    ImageRun,
    Table,
    TableCell,
    TableRow,
    Packer,
    Media,
    TextRun,
    PageBreak,
    AlignmentType,
    VerticalPosition
} = require("docx");

const path = require("path");
const fs = require("fs");
const os = require("os");

function cmToPt(cm) {
    return cm * 0.3937 * 72;
}

function createImageCell(imagePath) {
    return new TableCell({
        children: [
            new Paragraph({
                children: [
                    new ImageRun({
                        data: fs.readFileSync(imagePath),
                        transformation: {
                            width: cmToPt(12.14), 
                            height: cmToPt(9.12), 
                        },
                    }),
                ],
                alignment: AlignmentType.CENTER,
            }),
        ],
        margins: { top: cmToPt(0.5), bottom: cmToPt(0.5), left: cmToPt(0.5), right: cmToPt(0.5) }, // margen de 0.5 cms
    });
}

function createEmptyCell() {
    return new TableCell({
        children: [
            new Paragraph({
                children: [],
                alignment: AlignmentType.CENTER,
            }),
        ],
        margins: { top: cmToPt(2), bottom: cmToPt(2), left: cmToPt(2), right: cmToPt(2) }, // margen de 0.5 cms
    });
}

function createImageTable(images) {
    const rows = [];
    for (let i = 0; i < images.length; i += 2) {
        rows.push(new TableRow({
            children: [createImageCell(images[i]), images[i + 1] ? createImageCell(images[i + 1]) : createEmptyCell()],
        }));
    }
    return new Table({ rows, alignment: AlignmentType.CENTER }); 
}


const sections = [];

const contactInfo = new Paragraph({
    children: [
        new TextRun("CORREDORA MISS PROPIEDADES"),
        new TextRun({
            text: "\nContacto: +56 9 3684 0456 / contacto@misspropiedades.cl / www.misspropiedades.cl", break: 1}
        ),
    ],
    alignment: AlignmentType.CENTER,
    floating: VerticalPosition
});

async function CreateWord(event, wordName, imagesFolder) {
    return new Promise(async (resolve, reject) => {
        try {

            //Build para mac
            const imagenesPath = path.join(os.homedir(), 'Desktop', imagesFolder);
            const logoPath = path.join(os.homedir(), 'Desktop', 'logo.png');


            // Lee las imágenes
            const wordImages = fs.readdirSync(imagenesPath).filter(file => {
                const ext = path.extname(file).toLowerCase();
                return !['.mp4', '.mov', '.avi', '.mkv'].includes(ext);
            });
            
            const totalImages = wordImages.length;
            let processedImages = 0;

            // Crear las secciones
            const sections = [];

            for (let i = 0; i < 5; i++) {
                sections.push({
                    properties: {
                        page: {
                            margin: {
                                top: 233.8, // 3.3 cm
                                right: 141.7, // 1.5 cm
                                bottom: 198.4, // 3.5 cm
                                left: 141.7, // 1.5 cm
                                gutter: 0, // 0 cm
                                gutterPosition: 'left',
                            },
                        },
                    },
                    children: [],
                });
            }

            for (let i = 0; i < wordImages.length; i += 6) {
                sections.push({
                    properties: {
                        page: {
                            margin: {
                                top: 233.8, // 3.3 cm
                                right: 141.7, // 1.5 cm
                                bottom: 198.4, // 3.5 cm
                                left: 141.7, // 1.5 cm
                                gutter: 0, // 0 cm
                                gutterPosition: 'left',
                            },
                        },
                    },
                    headers: {
                        default: new Header({
                            children: [
                                new Paragraph({
                                    children: [
                                        new ImageRun({
                                            data: fs.readFileSync(logoPath),
                                            transformation: {
                                                width: 150,
                                                height: 130,
                                            },
                                        }),
                                    ],
                                    alignment: AlignmentType.LEFT,
                                }),
                            ],
                        }),
                    },
                    children: [
                        new Paragraph({}), 
                        new Paragraph({
                            children: [
                                new TextRun("REGISTRO FOTOGRÁFICO"),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                        new Paragraph({}), 
                        new Paragraph({}), 
                        createImageTable(wordImages.slice(i, i + 6).map(image => path.join(imagenesPath, image))),
                    ],
                    footers: {
                        default: new Footer({
                            children: [new Paragraph("Contact Info")], // Asegúrate de definir `contactInfo` adecuadamente
                        }),
                    },
                });

                processedImages += 6;
                const progress = Math.min((processedImages / totalImages) * 100, 100);
                event.sender.send('progress', progress);
            }

            const doc = new Document({ sections });
            const buffer = await Packer.toBuffer(doc);
            const wordPath = path.join(os.homedir(), 'Desktop', wordName + ".docx");
            fs.writeFileSync(wordPath, buffer);

            resolve();
        } catch (error) {
            console.error('Error al crear el archivo Word:', error);
            reject(error);
        }
    });
}

    
module.exports = {CreateWord}
