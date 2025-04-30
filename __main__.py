from typing import List

from docx import Document

# Constants
OUTPUT_FILE = 'Skills_Inventory.docx'
HEADERS: List[ str ] = [ 'Skill (Technologies)', 'Personal Rating', 'Years of Exp.', 'Last Used' ]
SKILLS_DATA: List[ List[ str ] ] = [
        [ '.NET (C#)', 'Advanced User', '5', '2024' ],
        [ 'Linux', 'Functional User', '3', '2025' ],
        [ 'Docker', 'Expert User', '4', '2025' ],
        [ 'Software Development Principles & Practices', 'Contributor', '7', '2024' ],
        [ 'Kafka', 'Knows about', '1', '2023' ],
        [ 'Kubernetes', 'Advanced User', '3', '2025' ],
        [ 'GitOps', 'Functional User', '2', '2024' ],
        [ 'GitHub Workflows', 'Expert User', '4', '2025' ],
        [ 'Azure', 'Advanced User', '5', '2024' ],
        [ 'Terraform', 'Functional User', '2', '2025' ],
        [ 'Configuration Management', 'Knows about', '1', '2023' ],
        [ 'Monitoring / Observability Tools (OTEL, Grafana, Prometheus)', 'Advanced User', '3', '2025' ],
        [ 'Agile / Scrum', 'Expert User', '6', '2024' ],
        [ 'Mentoring / Knowledge Sharing', 'Contributor', '5', '2025' ],
        [ 'CI/CD Pipeline Development', 'Advanced User', '4', '2024' ],
        [ 'Technical Documentation', 'Functional User', '3', '2025' ],
        [ 'Communication Skills', 'Expert User', '7', '2025' ],
        [ 'Problem Solving', 'Advanced User', '6', '2024' ],
        [ 'Task Prioritisation and Organisation', 'Functional User', '4', '2025' ],
        [ 'Flexibility & Adaptability to Change', 'Contributor', '5', '2025' ]
        ]


def create_skills_document( ) -> None :
    """
    Creates a Word document containing a skills inventory table.
    The document includes skills, ratings, experience, and last usage dates.
    
    Raises:
        Exception: If there's an error creating or saving the document
    """
    try :
        doc = Document( )
        doc.add_heading( 'Skills Inventory', 0 )

        table = doc.add_table( rows = 1, cols = len( HEADERS ) )
        table.style = 'Table Grid'
        table.autofit = True

        # Header row
        for i, header in enumerate( HEADERS ) :
            table.rows[ 0 ].cells[ i ].text = header

        # Data rows
        for skill_row in SKILLS_DATA :
            row_cells = table.add_row( ).cells
            for i, value in enumerate( skill_row ) :
                row_cells[ i ].text = value

        doc.save( OUTPUT_FILE )
        print( f"Document successfully created: {OUTPUT_FILE}" )

    except Exception as e :
        print( f"Error creating document: {str( e )}" )
        raise


if __name__ == '__main__' :
    try :
        create_skills_document( )
    except ImportError :
        print( "Error: The 'python-docx' library is not installed. Please install it using 'pip install python-docx'" )
    except Exception as e :
        print( f"An unexpected error occurred: {str( e )}" )
