import textwrap
import pandas as pd
from PIL import ImageFont, ImageDraw, Image
import os
import img2pdf

class CertificateGenerator:
    def __init__(self, 
                 template_single_line_path, 
                 template_two_line_path, 
                 excel_path, 
                 output_dir,
                 fonts,
                 colors,
                 max_line_width=95,
                 coordinates=None,
                 line_spacing=20):
        """
        Initialize Certificate Generator with templates and configuration
        """
        # Load templates
        self.template_single_line = Image.open(template_single_line_path)
        self.template_two_line = Image.open(template_two_line_path)
        
        # Load Excel data
        self.df = pd.read_excel(excel_path)
        
        # Create output directories if they don't exist
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(os.path.join(output_dir, 'png'), exist_ok=True)
        os.makedirs(os.path.join(output_dir, 'pdf'), exist_ok=True)
        
        # Store configuration
        self.output_dir = output_dir
        self.fonts = fonts
        self.colors = colors
        self.max_line_width = max_line_width
        self.line_spacing = line_spacing
        
        # Default coordinates if not provided
        self.coordinates = coordinates or {
            'paper_title_single_line': (938, 820),
            'paper_title_multi_line': (1017, 820),
            'participant_name': (1000, 620),
            'institute': (906, 765),
            'participant_type': (239, 777)
        }

    def _wrap_paper_title(self, paper_title):
        """
        Wrap paper title based on maximum line width
        
        Returns:
        - wrapped lines
        - template
        - boolean indicating if title is single-line
        """
        # If title is short, use single-line template
        if len(paper_title) <= self.max_line_width:
            return [paper_title], self.template_single_line, True
        
        # If title is long, wrap and use two-line template
        wrapped_lines = textwrap.wrap(paper_title, width=self.max_line_width)
        return wrapped_lines, self.template_two_line, False

    def _draw_text_centered(self, draw, text, font, x, y, color):
        """
        Draw text centered at given coordinates
        """
        # Calculate text bbox to center it
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        position = (x - text_width // 2, y)
        draw.text(position, text, font=font, fill=color)

    def generate_certificates(self):
        """
        Generate certificates for all participants in the Excel sheet
        """
        # Mapping of potential column names
        column_mapping = {
            'participant_name': ['Corresponding Author', 'Author Name', 'Name','Presenter Name'],
            'institute': ['Institute', 'University', 'College'],
            'participant_type': ['Participant Type', 'Type', 'Category'],
            'paper_title': ['Paper Title', 'Title', 'Research Title'],
            'paper_id': ['Paper ID', 'ID', 'Unique ID']
        }

        # Find correct column names
        def find_column(key):
            for possible_col in column_mapping[key]:
                if possible_col in self.df.columns:
                    return possible_col
            raise KeyError(f"Could not find a column for {key}")

        try:
            # Find the correct column names
            participant_name_col = find_column('participant_name')
            institute_col = find_column('institute')
            paper_title_col = find_column('paper_title')
            paper_id_col = find_column('paper_id')
            
            # Participant type is optional
            try:
                participant_type_col = find_column('participant_type')
            except KeyError:
                participant_type_col = None

        except KeyError as e:
            print(f"Error: {e}")
            print("Available columns:", list(self.df.columns))
            return

        for index, row in self.df.iterrows():
            # Extract participant details
            participant_name = row[participant_name_col]
            institute = row[institute_col]
            paper_title = row[paper_title_col]
            paper_id = row[paper_id_col]

            # Handle participant type (optional)
            participant_type = row[participant_type_col] if participant_type_col else "Participant"

            # Wrap paper title and select appropriate template
            wrapped_title, template, is_single_line = self._wrap_paper_title(str(paper_title))

            # Create a copy of the template
            image = template.copy()
            draw = ImageDraw.Draw(image)

            # Draw centered texts
            self._draw_text_centered(
                draw, 
                str(participant_name), 
                self.fonts['participant'], 
                self.coordinates['participant_name'][0], 
                self.coordinates['participant_name'][1], 
                self.colors['participant']
            )

            self._draw_text_centered(
                draw, 
                str(institute), 
                self.fonts['institute'], 
                self.coordinates['institute'][0], 
                self.coordinates['institute'][1], 
                self.colors['institute']
            )

            # Only draw participant type if it exists
            if participant_type_col:
                self._draw_text_centered(
                    draw, 
                    str(participant_type), 
                    self.fonts['participant_type'], 
                    self.coordinates['participant_type'][0], 
                    self.coordinates['participant_type'][1], 
                    self.colors['participant_type']
                )

            # Select coordinates based on title length
            title_coordinates = (
                self.coordinates['paper_title_single_line'] if is_single_line 
                else self.coordinates['paper_title_multi_line']
            )

            # Draw wrapped paper title
            current_y = title_coordinates[1]
            for line in wrapped_title:
                self._draw_text_centered(
                    draw, 
                    line, 
                    self.fonts['paper_title'], 
                    title_coordinates[0], 
                    current_y, 
                    self.colors['paper_title']
                )
                # Use font's font size instead of deprecated getsize method
                current_y += self.fonts['paper_title'].size + self.line_spacing

            # Save the PNG certificate
            png_output_path = os.path.join(self.output_dir, 'png', f"{paper_id}.png")
            pdf_output_path = os.path.join(self.output_dir, 'pdf', f"{paper_id}.pdf")
            
            # Save PNG
            image.save(png_output_path)
            
            # Convert PNG to PDF
            try:
                with open(pdf_output_path, "wb") as pdf_file:
                    pdf_file.write(img2pdf.convert(png_output_path))
                print(f"Certificate generated for: {paper_id} (PNG and PDF)")
            except Exception as e:
                print(f"Error converting {paper_id} to PDF: {e}")

# Main execution remains the same
def main():
    # Paths and configurations remain the same
    template_single_line_path = r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Assets\Templates\1\faculty certificate with sort title 5th.png"
    template_two_line_path = r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Assets\Templates\1\faculty certificate with large title 5th.png"
    excel_path = r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Assets\Excel Sheet\Book2(9).xlsx"
    output_dir = r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Gen_cert\New\21"

    # Fonts configuration remains the same
    fonts = {
        'participant': ImageFont.truetype(r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Assets\Fonts\Euphoria_Script\EuphoriaScript-Regular.ttf", 85),
        'institute': ImageFont.truetype(r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Assets\Fonts\Lora\Lora-VariableFont_wght.ttf", 35),
        'participant_type': ImageFont.truetype(r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Assets\Fonts\Lora\Lora-VariableFont_wght.ttf", 35),
        'paper_title': ImageFont.truetype(r"C:\Users\SANDILYA SUNDRAM\Desktop\Cert gen\Assets\Fonts\Lora\Lora-VariableFont_wght.ttf", 35)
    }

    # Colors configuration remains the same
    colors = {
        'participant': "#921717",
        'institute': "#4d0000",
        'participant_type': "#020202",
        'paper_title': "#4d0000"
    }

    # Updated coordinates configuration to include separate coordinates for single and multi-line titles
    coordinates = {
        'participant_name': (1000, 620),
        'institute': (906, 765),
        'participant_type': (263, 777),
        'paper_title_single_line': (938, 820),  # Coordinates for single-line titles
        'paper_title_multi_line': (1017, 820)    # Coordinates for multi-line titles
    }

    # Create certificate generator
    generator = CertificateGenerator(
        template_single_line_path,
        template_two_line_path,
        excel_path,
        output_dir,
        fonts,
        colors,
        max_line_width=95,  # Maximum characters per line for paper title
        coordinates=coordinates,
        line_spacing=20
    )

    # Generate certificates
    generator.generate_certificates()

if __name__ == "__main__":
    main()