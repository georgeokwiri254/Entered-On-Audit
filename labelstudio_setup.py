#!/usr/bin/env python3
"""
LabelStudio Setup for NER Annotation
Sets up LabelStudio project for reservation email NER annotation
"""

import json
import os
import requests
from pathlib import Path
from typing import List, Dict, Any
import subprocess
import time

class LabelStudioNERSetup:
    """Set up LabelStudio for NER annotation"""
    
    def __init__(self, 
                 labelstudio_url: str = "http://localhost:8080",
                 username: str = "admin@example.com",
                 password: str = "password123"):
        self.base_url = labelstudio_url.rstrip('/')
        self.username = username
        self.password = password
        self.api_key = None
        self.project_id = None
        
        # Entity labels for reservation emails
        self.entity_labels = [
            'MAIL_FIRST_NAME', 'MAIL_FULL_NAME', 'MAIL_ARRIVAL', 'MAIL_DEPARTURE',
            'MAIL_NIGHTS', 'MAIL_PERSONS', 'MAIL_ROOM', 'MAIL_RATE_CODE', 
            'MAIL_C_T_S', 'MAIL_NET_TOTAL', 'MAIL_TOTAL', 'MAIL_TDF', 
            'MAIL_ADR', 'MAIL_AMOUNT'
        ]
        
        # Entity colors for visualization
        self.entity_colors = {
            'MAIL_FIRST_NAME': '#FF6B6B',
            'MAIL_FULL_NAME': '#4ECDC4', 
            'MAIL_ARRIVAL': '#45B7D1',
            'MAIL_DEPARTURE': '#96CEB4',
            'MAIL_NIGHTS': '#FECA57',
            'MAIL_PERSONS': '#FF9FF3',
            'MAIL_ROOM': '#54A0FF',
            'MAIL_RATE_CODE': '#5F27CD',
            'MAIL_C_T_S': '#00D2D3',
            'MAIL_NET_TOTAL': '#FF6B35',
            'MAIL_TOTAL': '#2ED573',
            'MAIL_TDF': '#FFA502',
            'MAIL_ADR': '#3742FA',
            'MAIL_AMOUNT': '#2F3542'
        }
    
    def install_labelstudio(self):
        """Install LabelStudio via pip"""
        print("📦 Installing LabelStudio...")
        
        try:
            subprocess.check_call(["pip", "install", "label-studio", "label-studio-sdk"])
            print("✅ LabelStudio installed successfully")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install LabelStudio: {e}")
            return False
    
    def start_labelstudio_server(self, port: int = 8080):
        """Start LabelStudio server"""
        print(f"🚀 Starting LabelStudio server on port {port}...")
        print("Note: This will start the server in the background")
        
        # Create startup script
        startup_script = f"""#!/bin/bash
export LABEL_STUDIO_USERNAME={self.username}
export LABEL_STUDIO_PASSWORD={self.password}
label-studio start --host 0.0.0.0 --port {port} --data-dir ./labelstudio_data
"""
        
        with open("start_labelstudio.sh", "w") as f:
            f.write(startup_script)
        
        os.chmod("start_labelstudio.sh", 0o755)
        
        print("🔧 Created startup script: start_labelstudio.sh")
        print(f"Run: ./start_labelstudio.sh")
        print(f"Then access: http://localhost:{port}")
        print(f"Username: {self.username}")
        print(f"Password: {self.password}")
        
        return True
    
    def get_api_key(self) -> str:
        """Get API key from LabelStudio"""
        print("🔑 Getting API key...")
        
        # Login to get token
        login_url = f"{self.base_url}/user/login"
        login_data = {
            "username": self.username,
            "password": self.password
        }
        
        try:
            response = requests.post(login_url, json=login_data)
            if response.status_code == 200:
                # In newer versions, check for token in response
                token_data = response.json()
                if 'token' in token_data:
                    self.api_key = token_data['token']
                    print("✅ Got API key from login")
                    return self.api_key
            
            # Fallback: try to get from account settings
            # This requires manual setup in the LabelStudio UI
            print("⚠️  Could not get API key automatically")
            print("Please:")
            print("1. Go to http://localhost:8080")
            print("2. Login with your credentials")
            print("3. Go to Account & Settings > Access Token")
            print("4. Generate a new token")
            print("5. Copy the token and use it in the setup")
            
            return None
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Error connecting to LabelStudio: {e}")
            print("Make sure LabelStudio is running on the specified URL")
            return None
    
    def create_labeling_config(self) -> str:
        """Create LabelStudio labeling configuration for NER"""
        
        # Build entity choices
        choices = []
        for i, entity in enumerate(self.entity_labels):
            color = self.entity_colors.get(entity, '#FF6B6B')
            choices.append(f'<Choice value="{entity}" background="{color}"/>')
        
        choices_xml = '\n      '.join(choices)
        
        # LabelStudio XML configuration for NER
        config = f'''
<View>
  <Text name="text" value="$text"/>
  <Labels name="label" toName="text">
    <Label value="MAIL_FIRST_NAME" background="#FF6B6B"/>
    <Label value="MAIL_FULL_NAME" background="#4ECDC4"/>
    <Label value="MAIL_ARRIVAL" background="#45B7D1"/>
    <Label value="MAIL_DEPARTURE" background="#96CEB4"/>
    <Label value="MAIL_NIGHTS" background="#FECA57"/>
    <Label value="MAIL_PERSONS" background="#FF9FF3"/>
    <Label value="MAIL_ROOM" background="#54A0FF"/>
    <Label value="MAIL_RATE_CODE" background="#5F27CD"/>
    <Label value="MAIL_C_T_S" background="#00D2D3"/>
    <Label value="MAIL_NET_TOTAL" background="#FF6B35"/>
    <Label value="MAIL_TOTAL" background="#2ED573"/>
    <Label value="MAIL_TDF" background="#FFA502"/>
    <Label value="MAIL_ADR" background="#3742FA"/>
    <Label value="MAIL_AMOUNT" background="#2F3542"/>
  </Labels>
</View>
'''
        
        return config.strip()
    
    def create_project(self, project_name: str = "Reservation Email NER") -> bool:
        """Create LabelStudio project"""
        if not self.api_key:
            print("❌ No API key available. Cannot create project.")
            return False
        
        print(f"🏗️  Creating project: {project_name}")
        
        headers = {
            'Authorization': f'Token {self.api_key}',
            'Content-Type': 'application/json'
        }
        
        labeling_config = self.create_labeling_config()
        
        project_data = {
            "title": project_name,
            "description": "Named Entity Recognition for reservation email extraction",
            "label_config": labeling_config,
            "expert_instruction": """
            Please annotate the following entities in reservation emails:
            
            📧 GUEST INFORMATION:
            • MAIL_FIRST_NAME: Guest first name
            • MAIL_FULL_NAME: Guest last name or full name
            
            📅 DATES & DURATION:
            • MAIL_ARRIVAL: Check-in date
            • MAIL_DEPARTURE: Check-out date  
            • MAIL_NIGHTS: Number of nights
            
            🏨 ACCOMMODATION:
            • MAIL_PERSONS: Number of persons/adults
            • MAIL_ROOM: Room type or room code
            • MAIL_RATE_CODE: Rate/promo code
            
            🏢 AGENCY & FINANCIAL:
            • MAIL_C_T_S: Travel agency/company name
            • MAIL_NET_TOTAL: Net total amount
            • MAIL_TOTAL: Total amount (including taxes)
            • MAIL_TDF: Tourism tax/fees
            • MAIL_ADR: Average daily rate
            • MAIL_AMOUNT: Base amount
            
            Guidelines:
            - Select the most specific text for each entity
            - Avoid overlapping annotations
            - Include currency symbols with amounts
            - Mark dates in the format they appear
            """,
            "show_instruction": True,
            "show_skip_button": True,
            "enable_empty_annotation": False
        }
        
        try:
            create_url = f"{self.base_url}/api/projects"
            response = requests.post(create_url, json=project_data, headers=headers)
            
            if response.status_code == 201:
                project_info = response.json()
                self.project_id = project_info['id']
                print(f"✅ Created project with ID: {self.project_id}")
                return True
            else:
                print(f"❌ Failed to create project: {response.status_code}")
                print(response.text)
                return False
                
        except requests.exceptions.RequestException as e:
            print(f"❌ Error creating project: {e}")
            return False
    
    def convert_bio_to_labelstudio(self, bio_data_path: str, output_path: str = "labelstudio_tasks.json"):
        """Convert BIO format data to LabelStudio tasks"""
        print(f"🔄 Converting BIO data to LabelStudio format...")
        
        # Load BIO data
        with open(bio_data_path, 'r', encoding='utf-8') as f:
            bio_data = json.load(f)
        
        labelstudio_tasks = []
        
        for record in bio_data:
            # Reconstruct text from tokens
            text = " ".join(record['tokens'])
            
            # Create task
            task = {
                "data": {
                    "text": text,
                    "email_id": record['email_id'],
                    "agency": record['agency']
                },
                "predictions": []  # We'll add weak labels as predictions
            }
            
            # Convert BIO labels to LabelStudio annotations
            predictions = self.bio_to_labelstudio_annotations(record['tokens'], record['labels'])
            if predictions:
                task["predictions"] = [{"result": predictions}]
            
            labelstudio_tasks.append(task)
        
        # Save tasks
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(labelstudio_tasks, f, indent=2, ensure_ascii=False)
        
        print(f"💾 Saved {len(labelstudio_tasks)} tasks to {output_path}")
        return output_path
    
    def bio_to_labelstudio_annotations(self, tokens: List[str], labels: List[str]) -> List[Dict]:
        """Convert BIO labels to LabelStudio annotations"""
        annotations = []
        current_entity = None
        current_start = 0
        current_text = ""
        
        text_offset = 0
        
        for i, (token, label) in enumerate(zip(tokens, labels)):
            if label.startswith('B-'):
                # End previous entity if exists
                if current_entity:
                    annotations.append({
                        "from_name": "label",
                        "to_name": "text", 
                        "type": "labels",
                        "value": {
                            "start": current_start,
                            "end": text_offset,
                            "text": current_text.strip(),
                            "labels": [current_entity]
                        }
                    })
                
                # Start new entity
                current_entity = label[2:]  # Remove B- prefix
                current_start = text_offset
                current_text = token
                
            elif label.startswith('I-') and current_entity:
                # Continue current entity
                current_text += " " + token
                
            else:
                # End current entity if exists
                if current_entity:
                    annotations.append({
                        "from_name": "label",
                        "to_name": "text",
                        "type": "labels", 
                        "value": {
                            "start": current_start,
                            "end": text_offset,
                            "text": current_text.strip(),
                            "labels": [current_entity]
                        }
                    })
                    current_entity = None
            
            # Update text offset
            text_offset += len(token) + 1  # +1 for space
        
        # Handle final entity
        if current_entity:
            annotations.append({
                "from_name": "label", 
                "to_name": "text",
                "type": "labels",
                "value": {
                    "start": current_start,
                    "end": text_offset - 1,
                    "text": current_text.strip(),
                    "labels": [current_entity]
                }
            })
        
        return annotations
    
    def import_tasks(self, tasks_file: str) -> bool:
        """Import tasks to LabelStudio project"""
        if not self.api_key or not self.project_id:
            print("❌ Need API key and project ID to import tasks")
            return False
        
        print(f"📥 Importing tasks from {tasks_file}...")
        
        headers = {
            'Authorization': f'Token {self.api_key}',
        }
        
        try:
            import_url = f"{self.base_url}/api/projects/{self.project_id}/import"
            
            with open(tasks_file, 'rb') as f:
                files = {'file': f}
                response = requests.post(import_url, files=files, headers=headers)
            
            if response.status_code == 201:
                result = response.json()
                print(f"✅ Imported {result.get('task_count', 'unknown')} tasks")
                return True
            else:
                print(f"❌ Failed to import tasks: {response.status_code}")
                print(response.text)
                return False
                
        except requests.exceptions.RequestException as e:
            print(f"❌ Error importing tasks: {e}")
            return False
    
    def generate_setup_instructions(self):
        """Generate complete setup instructions"""
        instructions = f"""
# LabelStudio NER Setup Instructions

## 1. Installation & Startup
```bash
# Install LabelStudio
pip install label-studio label-studio-sdk

# Start server
export LABEL_STUDIO_USERNAME={self.username}
export LABEL_STUDIO_PASSWORD={self.password}
label-studio start --host 0.0.0.0 --port 8080 --data-dir ./labelstudio_data

# Or use the generated script
./start_labelstudio.sh
```

## 2. Access LabelStudio
- URL: http://localhost:8080
- Username: {self.username}
- Password: {self.password}

## 3. Get API Token
1. Login to LabelStudio
2. Go to Account & Settings > Access Token
3. Generate new token
4. Copy token for programmatic access

## 4. Entity Labels
The following entities should be annotated:

**Guest Information:**
- MAIL_FIRST_NAME: Guest first name
- MAIL_FULL_NAME: Guest last name

**Dates & Duration:**
- MAIL_ARRIVAL: Check-in date
- MAIL_DEPARTURE: Check-out date
- MAIL_NIGHTS: Number of nights

**Accommodation:**
- MAIL_PERSONS: Number of persons
- MAIL_ROOM: Room type/code
- MAIL_RATE_CODE: Rate/promo code

**Financial & Agency:**
- MAIL_C_T_S: Travel agency name
- MAIL_NET_TOTAL: Net total amount
- MAIL_TOTAL: Total with taxes
- MAIL_TDF: Tourism tax/fees
- MAIL_ADR: Average daily rate
- MAIL_AMOUNT: Base amount

## 5. Annotation Guidelines
- Select precise text spans for each entity
- Include currency symbols with amounts
- Mark dates in their original format
- Avoid overlapping annotations
- Focus on accuracy over speed

## 6. Export Annotations
After annotation, export in JSON format:
1. Go to project > Export
2. Select JSON format
3. Download annotations
4. Convert back to BIO format for training

## 7. Quality Control
- Review annotations for consistency
- Check entity boundaries
- Validate against original parser outputs
- Resolve conflicts between annotators
"""
        
        # Save instructions
        with open("labelstudio_setup_instructions.md", "w") as f:
            f.write(instructions)
        
        print("📋 Generated setup instructions: labelstudio_setup_instructions.md")
        return instructions

def main():
    """Main setup workflow"""
    print("🏷️  LabelStudio NER Setup")
    print("=" * 50)
    
    # Create setup instance
    setup = LabelStudioNERSetup()
    
    # Install LabelStudio
    if not setup.install_labelstudio():
        print("❌ Installation failed. Please install manually:")
        print("pip install label-studio label-studio-sdk")
        return
    
    # Generate startup script
    setup.start_labelstudio_server()
    
    # Generate setup instructions
    instructions = setup.generate_setup_instructions()
    print("\n📋 Setup Instructions Generated")
    print("=" * 50)
    print(instructions)
    
    # Save labeling configuration
    config = setup.create_labeling_config()
    with open("labelstudio_ner_config.xml", "w") as f:
        f.write(config)
    print("💾 Saved labeling config: labelstudio_ner_config.xml")
    
    print("\n✅ LabelStudio setup complete!")
    print("\nNext steps:")
    print("1. Start LabelStudio: ./start_labelstudio.sh")
    print("2. Access http://localhost:8080")
    print("3. Create project with the provided config")
    print("4. Import your training data")
    print("5. Begin annotation!")

if __name__ == "__main__":
    main()