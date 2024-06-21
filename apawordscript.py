import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def create_apa_document(title, author, institution, course, instructor, due_date, abstract, keywords, content, references, is_professional=False):
    doc = docx.Document()
    
    # Set up the document
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Set font for entire document
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Enable double spacing
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(0)
    paragraph_format.line_spacing = 2.0

    # Title Page
    if is_professional:
        running_head = title.upper()[:50]  # Limit to 50 characters
        header = section.header
        header_para = header.paragraphs[0]
        header_para.text = f"Running head: {running_head}"
        header_para.style = doc.styles['Header']
        
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title)
    title_run.bold = True

    for item in [author, institution]:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.add_run(item)

    if not is_professional:
        for item in [course, instructor, due_date]:
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.add_run(item)

    doc.add_page_break()

    # Abstract (only for professional papers or if specifically requested)
    if is_professional or abstract:
        doc.add_heading('Abstract', level=1)
        doc.add_paragraph(abstract)
        
        # Keywords
        keywords_para = doc.add_paragraph('Keywords: ')
        keywords_para.add_run(', '.join(keywords)).italic = True

        doc.add_page_break()

    # Content
    for paragraph in content.split('\n\n'):
        p = doc.add_paragraph(paragraph)
        p.paragraph_format.first_line_indent = Inches(0.5)

    doc.add_page_break()

    # References
    references_heading = doc.add_paragraph("References")
    references_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    references_heading.runs[0].bold = True

    for ref in references:
        para = doc.add_paragraph(ref)
        para.paragraph_format.left_indent = Inches(0.5)
        para.paragraph_format.first_line_indent = Inches(-0.5)

    # Add page numbers
    add_page_number(doc.sections[0])

    return doc

def add_page_number(section):
    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    field = OxmlElement('w:fldSimple')
    field.set(qn('w:instr'), 'PAGE')
    run = paragraph.add_run()
    run._element.append(field)
    
    return paragraph

# Example usage
title = "Giotto di Bondone and the Development of Renaissance Painting: Narrative Tradition and Naturalism as Didactic Tools"
author = "Tyler B Gibbs"
institution = "University of Oklahoma"
course = "Course Number: LSTD 3173-501"
instructor = "Allison Palmer"
due_date = "June 17, 2024"

abstract = """This essay explores the significant contributions of Giotto di Bondone to the development of Renaissance painting, focusing on his innovative narrative techniques and naturalistic style. As a pivotal figure in the Early Renaissance of Italy, Giotto revolutionized the way religious stories were depicted in art, moving away from the static Byzantine tradition towards a more dynamic and emotionally expressive approach. This paper examines how Giotto's narrative tradition served as an effective didactic tool for church congregations and how his groundbreaking naturalism enhanced viewers' understanding of his work. By analyzing key works such as the Scrovegni Chapel frescoes and the Ognissanti Madonna, this essay demonstrates how Giotto's innovations in spatial depth, humanization of religious figures, and emotional expressiveness laid the foundation for the artistic developments that would characterize the Renaissance. The lasting impact of Giotto's work on subsequent generations of artists and its role in shaping the course of Western art history are also discussed."""

keywords = ["Giotto di Bondone", "Renaissance painting", "narrative tradition", "naturalism", "didactic art"]

content = """
Giotto di Bondone and the Development of Renaissance Painting: Narrative Tradition and Naturalism as Didactic Tools

Giotto di Bondone, a pivotal figure in the Early Renaissance of Italy, played a crucial role in developing what came to be known as the "Renaissance" style of painting. His innovative approach to art, particularly in the realm of religious narrative painting, marked a significant departure from the Byzantine style that had dominated European art for centuries. This essay will explore how Giotto's narrative tradition served as a didactic tool for church congregations and how his groundbreaking naturalism enhanced viewers' understanding of his work.

The Narrative Tradition in Giotto's Work

Giotto's approach to narrative painting revolutionized the way religious stories were depicted and understood by viewers. Unlike the static, iconic representations of the Byzantine tradition, Giotto's paintings told stories through dynamic compositions and emotionally expressive figures. This narrative approach was particularly well-suited to the didactic needs of the Church, as it allowed for a more engaging and accessible presentation of biblical stories and religious teachings.

One of the most notable examples of Giotto's narrative style can be seen in his frescoes in the Scrovegni Chapel in Padua. Completed around 1305, this cycle of paintings depicts scenes from the lives of the Virgin Mary and Christ. Each scene is carefully composed to convey the key elements of the story, with figures arranged in a way that guides the viewer's eye through the narrative.

For instance, in the "Lamentation" fresco, Giotto arranges the figures around the body of Christ in a circular composition, creating a sense of movement and emotional intensity. The grieving figures' poses and expressions convey the story's emotional core, allowing viewers to connect with the narrative on a deeper level. This approach made the religious stories more relatable and memorable for the congregation, serving as an effective teaching tool.

Giotto's New Naturalism

One of Giotto's most significant contributions to the development of Renaissance painting was his introduction of greater naturalism. This shift towards more lifelike representations was a marked departure from the stylized figures of Byzantine art. Giotto's naturalism served not only to make his paintings more visually appealing but also to enhance their effectiveness as didactic tools.

Spatial Depth and Perspective

Giotto's pioneering use of spatial depth and rudimentary perspective helped create a more immersive visual experience for viewers. While not yet employing the mathematical perspective that would be developed later in the Renaissance, Giotto's paintings show a clear attempt to create the illusion of three-dimensional space on a two-dimensional surface.

In works like the "Ognissanti Madonna" (c. 1310), Giotto depicts the throne on which the Virgin Mary sits as a three-dimensional structure, receding into space. This spatial awareness helps to ground the figures in a more relatable, physical reality, making the scene feel more immediate and accessible to viewers (Zucker & Harris, 2020). The sense of depth and volume in Giotto's paintings was revolutionary for its time, setting the stage for the further development of perspective in Renaissance art.

Humanization of Religious Figures

Another key aspect of Giotto's naturalism was his humanization of religious figures. Rather than depicting saints and biblical characters as remote, otherworldly beings, Giotto presented them with recognizable human emotions and physical characteristics. This approach made the religious narratives more relatable and understandable to the average viewer.

In the "Ognissanti Madonna," for example, the Christ Child is portrayed not as a miniature adult, as was common in Byzantine art, but as a more realistic infant. The child's pose and interaction with Mary create a sense of tenderness and humanity that viewers could easily connect with their own experiences (Zucker & Harris, 2020). This humanization of religious figures was a significant departure from earlier artistic traditions and played a crucial role in making religious art more accessible to a broader audience.

Emotional Expressiveness

Giotto's figures are notable for their emotional expressiveness, a quality that was largely absent in earlier medieval art. By depicting characters with recognizable human emotions, Giotto allowed viewers to empathize more deeply with the stories being told.

In scenes such as the "Betrayal of Christ" in the Scrovegni Chapel, the facial expressions and body language of the figures convey complex emotions like betrayal, sorrow, and anger. This emotional realism helped viewers to engage more fully with the narrative, enhancing their understanding and retention of the religious teachings being conveyed. Giotto's ability to capture and convey human emotions in his paintings was groundbreaking for its time and became a hallmark of Renaissance art.

The Role of Light and Color

Giotto's use of light and color also contributed to the naturalism and didactic effectiveness of his paintings. Unlike the flat, gold-dominated palette of Byzantine art, Giotto employed a more varied and subtle range of colors to create a sense of depth and volume. His understanding of how light affects color and form allowed him to model figures and objects more convincingly, further enhancing the realism of his scenes.

In the Scrovegni Chapel frescoes, for example, Giotto uses light and shadow to create a sense of three-dimensionality in the figures and architectural elements. This not only made the scenes more visually engaging but also helped to highlight important elements of the narrative, guiding the viewer's attention and enhancing their understanding of the story being told.

Giotto's Influence on Renaissance Art

Giotto's innovations in narrative painting and naturalism laid the groundwork for the artistic developments that would characterize the Renaissance. His approach to storytelling through visual means and his emphasis on human emotion and physical reality became fundamental principles of Renaissance art.

Later artists, such as Masaccio and Fra Angelico, built upon Giotto's foundations, further developing techniques of perspective and naturalism. The narrative traditions established by Giotto continued to be refined and expanded throughout the Renaissance, culminating in the grand narrative cycles of artists like Michelangelo and Raphael.

The Impact on Religious Education

Giotto's new approach to religious art had a profound impact on religious education in the late medieval and early Renaissance periods. By making biblical stories and religious teachings more accessible and engaging through his naturalistic style and narrative techniques, Giotto helped to bridge the gap between the Church and its congregation.

The increased realism and emotional resonance of Giotto's paintings allowed viewers to connect more deeply with religious subjects, potentially strengthening their faith and understanding of Christian doctrine. This didactic function of art became increasingly important in the following centuries, with religious institutions commissioning artworks specifically designed to educate and inspire the faithful.

Giotto's Legacy in Western Art

The influence of Giotto's innovations extended far beyond the Renaissance period. His emphasis on naturalism and emotional expressiveness laid the foundation for the development of Western art as a whole. The idea that art should strive to represent the world as it appears to the human eye, rather than adhering to stylized conventions, became a central tenet of Western artistic tradition.

Moreover, Giotto's narrative approach to painting, which emphasized storytelling and emotional engagement, influenced the development of history painting as a genre. This genre, which depicted significant historical, mythological, or religious events, remained a dominant form of artistic expression well into the 19th century.

Conclusion

Giotto di Bondone's contributions to the development of Renaissance painting were profound and far-reaching. His narrative approach to religious art served as an effective didactic tool for church congregations, making complex theological concepts more accessible and engaging. The new naturalism he introduced, characterized by spatial depth, humanized religious figures, and emotional expressiveness, enhanced viewers' understanding and connection to the works.

By bridging the gap between the stylized forms of Byzantine art and the more realistic representations of the High Renaissance, Giotto played a crucial role in shaping the course of Western art history. His innovations not only transformed the visual language of painting but also redefined the relationship between art and viewer, paving the way for the artistic and cultural flowering of the Renaissance.

Giotto's legacy continues to influence our understanding of art's role in society, particularly its power to educate, inspire, and evoke emotion. His work reminds us that art is not merely a decorative or aesthetic pursuit, but a powerful tool for communication and human connection. As we continue to study and appreciate Giotto's contributions, we gain insight not only into the development of Renaissance art but also into the enduring power of visual storytelling in shaping our cultural and spiritual experiences."""

references = [
    "Adams, L. S. (2013). Italian Renaissance Art. Westview Press.",
    "Labatt, A., & Appleyard, C. (2004). Mendicant Orders in the Medieval World. The Metropolitan Museum of Art. http://www.metmuseum.org/toah/hd/mend/hd_mend.htm",
    "Zucker, S., & Harris, B. (2015, December 11). Cimabue, Santa Trinita Madonna and Child Enthroned. In Smarthistory. Retrieved June 10, 2024, from https://smarthistory.org/cimabue-santa-trinita-madonna",
    "Zucker, S., & Harris, B. (2020, November 23). Giotto, The *Ognissanti Madonna* and *Child Enthroned*. In Smarthistory. Retrieved June 10, 2024, from https://smarthistory.org/giotto-the-ognissanti-madonna/"
]

# Set is_professional to True for professional papers, False for student papers
is_professional = False

doc = create_apa_document(title, author, institution, course, instructor, due_date, abstract, keywords, content, references, is_professional)
doc.save('Giotto_Renaissance_Painting_Essay.docx')