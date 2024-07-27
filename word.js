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
                            width: cmToPt(9.12), 
                            height: cmToPt(6.83), 
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
        new TextRun("Contacto: +56 9 3684 0456 / contacto@misspropiedades.cl / www.misspropiedades.cl"),
    ],
    alignment: AlignmentType.CENTER,
    floating: VerticalPosition
});

async function CreateWord(event, wordImages, wordName){
    return new Promise (async (resolve, reject) => {

        try {

            const totalImages = wordImages.length;
            let processedImages = 0;

            for (let i = 0; i < wordImages.length; i += 6) {
                sections.push({
                    children: [
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    //For local testing
                                    data: fs.readFileSync("logo.png"),
                                    //data: fs.readFileSync("./resources/app/logo.png"),
                                    transformation: {
                                        width: 180, 
                                        height: 100, 
                                    },
                                }),
                            ],
                            alignment: AlignmentType.LEFT,
                        }),
                        new Paragraph({}), 
                        new Paragraph({}), 
                        new Paragraph({
                            children: [
                                new TextRun("REGISTRO FOTOGRÁFICO"),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                        new Paragraph({}), 
                        new Paragraph({}), 
                        //Para local
                         createImageTable(wordImages.slice(i, i + 6).map(image => path.join("./images", image))),
                        //createImageTable(wordImages.slice(i, i + 6).map(image => path.join("./resources/app/images", image))),
                    ],
                    footers: {
                        default: new Footer({
                            children: [contactInfo],
                        }),
                    },
                });

                processedImages += 6;
                const progress = Math.min((processedImages / totalImages) * 100, 100);
                event.sender.send('progress', progress);

            }
    
            const doc = new Document({ sections });
    
            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(wordName + ".docx", buffer);

            resolve();
        } catch (error) {
            console.error('Error al crear el archivo Word:', error);
            reject(error);
        }
    });
}
    
module.exports = {CreateWord}