import os
import pypandoc

input_dir = './Mayank'  # Change to your target directory
output_dir = './converted'  # Output folder

os.makedirs(output_dir, exist_ok=True)

for filename in os.listdir(input_dir):
    if filename.endswith('.odt'):
        input_path = os.path.join(input_dir, filename)
        output_filename = os.path.splitext(filename)[0] + '.docx'
        output_path = os.path.join(output_dir, output_filename)

        try:
            print(f"Converting: {filename}")
            pypandoc.convert_file(input_path, 'docx', outputfile=output_path)
            print(f"Saved to: {output_path}")
        except Exception as e:
            print(f"Failed to convert {filename}: {e}")

print("Batch conversion complete.")
