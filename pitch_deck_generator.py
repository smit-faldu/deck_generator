import json
import re
from typing import List, Dict, Any, Optional
from pptx import Presentation
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.agents import Tool, AgentExecutor, create_react_agent
from langchain.prompts import PromptTemplate
from langchain.tools import BaseTool
from langchain_core.messages import HumanMessage, AIMessage
from langchain.memory import ConversationBufferMemory
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")

# Initialize the model
model = ChatGoogleGenerativeAI(
    model="gemini-1.5-pro",
    temperature=0.2,
    google_api_key=GOOGLE_API_KEY,
    convert_system_message_to_human=True
)

class PowerPointStructureTool(BaseTool):
    """Tool for extracting structure from PowerPoint template"""
    name: str = "extract_ppt_structure"
    description: str = "Extract the structure and word counts from a PowerPoint template file"
    memory: Optional[ConversationBufferMemory] = None
    last_structure: Optional[List] = None
    
    def __init__(self, **data):
        super().__init__(**data)
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True,
            output_key="output"
        )
        self.last_structure = None
    
    def _run(self, file_path: str) -> Dict:
        try:
            # Handle case where file_path is a JSON string
            if isinstance(file_path, str) and file_path.startswith('{'):
                try:
                    file_path_dict = json.loads(file_path)
                    file_path = file_path_dict.get('template_path', file_path)
                except json.JSONDecodeError:
                    pass
            
            # Ensure the file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"PowerPoint template not found at: {file_path}")
            
            prs = Presentation(file_path)
            structure = []
            
            for idx, slide in enumerate(prs.slides):
                slide_info = {
                    'index': idx,
                    'placeholders': [],
                    'word_need': []
                }
                
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.has_text_frame:
                        text = shape.text.strip()
                        if text:
                            slide_info['placeholders'].append(text)
                            word_count = len(text.split())
                            slide_info['word_need'].append(word_count)
                
                if slide_info['placeholders']:
                    structure.append(slide_info)
            
            # Store structure in memory
            self.last_structure = structure
            
            # Store the interaction in memory
            self.memory.save_context(
                {"input": f"Extract structure from template: {file_path}"},
                {"output": structure}
            )
            
            return {"structure": structure, "status": "success"}
            
        except Exception as e:
            return {"error": f"Failed to extract PowerPoint structure: {str(e)}", "status": "error"}
    
    def _arun(self, file_path: str):
        raise NotImplementedError("Async version not implemented")

class StructureAnalysisTool(BaseTool):
    """Tool for analyzing PowerPoint structure"""
    name: str = "analyze_structure"
    description: str = "Analyze the PowerPoint structure to determine slide purposes and requirements"
    memory: Optional[ConversationBufferMemory] = None
    last_analysis: Optional[List] = None
    
    def __init__(self, **data):
        super().__init__(**data)
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True,
            output_key="output"
        )
        self.last_analysis = None
    
    def _run(self, structure: List) -> Dict:
        try:
            prompt = f"""
            Analyze this PowerPoint structure and determine the purpose of each slide.
            For each slide, identify:
            1. The main purpose (e.g., title, introduction, problem statement, etc.)
            2. The type of content needed
            3. The tone and style required
            4. Key elements that should be included

            Structure to analyze:
            {structure}

            For each slide, provide the analysis in this format:
            {{
                index: [slide number],
                purpose: [main purpose],
                content: [type of content needed],
                tone: [required tone and style],
                key_elements: [list of key elements],
                word_required: [list of word counts]
            }}

            Example:
            {{
                index: 0,
                purpose: "Title slide introducing the company",
                content: "Header and contact information",
                tone: "Professional and welcoming",
                key_elements: ["Company name", "Tagline", "Contact details"],
                word_required: [3, 2, 5]
            }}
            """
            
            response = model.invoke(prompt)
            analysis = response.content
            
            result = {"analysis": analysis, "status": "success"}
            
            # Store analysis in memory
            self.last_analysis = analysis
            
            # Store the interaction in memory
            self.memory.save_context(
                {"input": "Analyze PowerPoint structure"},
                {"output": analysis}
            )
            
            return result
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _arun(self, structure: List):
        raise NotImplementedError("This tool does not support async")

class ContentGenerationTool(BaseTool):
    """Tool for generating pitch deck content"""
    name: str = "generate_content"
    description: str = "Generate content for the pitch deck based on topic"
    memory: Optional[ConversationBufferMemory] = None
    last_content: Optional[List] = None
    
    def __init__(self, **data):
        super().__init__(**data)
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True,
            output_key="output"
        )
        self.last_content = None
    
    def _run(self, topic: str, structure: List, analysis: str) -> Dict:
        try:
            prompt = f"""Generate content for a pitch deck about: {topic}

            Use this structure and analysis to guide your content generation:

            Structure:
            {structure}

            Analysis:
            {analysis}

            For each slide, generate content in this format:
            {{
                index: [slide number],
                content: [generated content for each placeholder],
                word_count: [actual word count for each content]
            }}

            Example:
            {{
                index: 0,
                content: ["TechFlow AI", "Revolutionizing Business Automation", "contact@techflow.ai"],
                word_count: [3, 2, 1]
            }}

            Requirements:
            1. Content must be professional and business-appropriate
            2. Follow the structure and analysis exactly
            3. Match word counts as specified
            4. Keep content concise and impactful
            5. Include specific examples and data points

            IMPORTANT: Return ONLY the JSON content without any markdown formatting or code block markers.
            """
            
            response = model.invoke(prompt)
            content = response.content
            
            # Clean up the response by removing markdown formatting
            content = content.replace('```json', '').replace('```', '').strip()
            
            result = {"content": content, "status": "success"}
            
            # Store content in memory
            self.last_content = content
            
            # Store the interaction in memory
            self.memory.save_context(
                {"input": f"Generate content for topic: {topic}"},
                {"output": content}
            )
            
            return result
        except Exception as e:
            return {"error": str(e), "status": "error"}
    
    def _arun(self, topic: str, structure: List, analysis: str):
        raise NotImplementedError("This tool does not support async")

class PowerPointCreatorTool(BaseTool):
    """Tool for creating the final PowerPoint presentation"""
    name: str = "create_powerpoint"
    description: str = "Create the final PowerPoint presentation with the generated content"
    memory: Optional[ConversationBufferMemory] = None
    content_tool: Optional[ContentGenerationTool] = None
    
    def __init__(self, **data):
        super().__init__(**data)
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True,
            output_key="output"
        )
        self.content_tool = None
    
    def set_content_tool(self, content_tool: ContentGenerationTool):
        """Set the content generation tool"""
        self.content_tool = content_tool
    
    def _analyze_template(self, template_path):
        """Analyze the template structure to understand its layout and content areas"""
        prs = Presentation(template_path)
        template_structure = []
        
        for idx, slide in enumerate(prs.slides):
            slide_info = {
                'index': idx,
                'placeholders': [],
                'shapes': []
            }
            
            # Analyze text placeholders
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.has_text_frame:
                    text = shape.text.strip()
                    if text:
                        slide_info['placeholders'].append(text)
                
                # Store shape information
                shape_info = {
                    'type': shape.shape_type,
                    'name': getattr(shape, 'name', ''),
                    'id': getattr(shape, 'shape_id', '')
                }
                slide_info['shapes'].append(shape_info)
            
            template_structure.append(slide_info)
        
        return template_structure
    
    def _run(self, template_path: str, output_path: str = "output/AI_Based_Pitch_Deck.pptx") -> Dict:
        try:
            if not self.content_tool or not self.content_tool.last_content:
                return {"error": "No content found. Please run generate_content first.", "status": "error"}
            
            content = self.content_tool.last_content
            
            # Parse content if it's a string
            if isinstance(content, str):
                try:
                    content = json.loads(content)
                except json.JSONDecodeError:
                    # Try to extract JSON from the string
                    json_match = re.search(r'\{[\s\S]*\}', content)
                    if json_match:
                        content = json.loads(json_match.group())
                    else:
                        return {"error": "Invalid content format", "status": "error"}
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Analyze the template structure
            template_structure = self._analyze_template(template_path)
            print("Template structure:", json.dumps(template_structure, indent=2))
            
            # Load the template
            prs = Presentation(template_path)
            
            # Get the slides data - handle both list and dict formats
            if isinstance(content, dict):
                slides_data = content.get('slides', [])
            else:
                slides_data = content  # Assume it's already a list of slides
            
            print("Content slides data:", json.dumps(slides_data, indent=2))
            
            # Text replacement using structure-based mapping
            for slide_data in slides_data:
                index = slide_data.get('index')
                if index is None or index >= len(prs.slides):
                    continue

                slide = prs.slides[index]
                
                # Handle different content formats
                if 'placeholders' in slide_data:
                    new_texts = slide_data['placeholders']
                elif 'content' in slide_data:
                    new_texts = slide_data['content']
                else:
                    print(f"Warning: No content found for slide {index}")
                    continue
                
                # Get the template structure for this slide
                template_slide = template_structure[index]
                original_placeholders = template_slide['placeholders']
                
                # Create mapping between original and new texts
                # If we have more new texts than placeholders, use only the first ones
                # If we have fewer new texts than placeholders, pad with empty strings
                max_items = max(len(original_placeholders), len(new_texts))
                padded_new_texts = new_texts + [''] * (max_items - len(new_texts))
                text_map = dict(zip(original_placeholders, padded_new_texts))
                
                print(f"Slide {index} mapping:", text_map)

                # Replace text in each shape
                for shape in slide.shapes:
                    if not hasattr(shape, "text") or not shape.has_text_frame:
                        continue

                    original_text = shape.text.strip()
                    if original_text in text_map:
                        new_text = text_map[original_text]

                        # Keep style from the first run
                        tf = shape.text_frame
                        if tf.paragraphs and tf.paragraphs[0].runs:
                            first_run = tf.paragraphs[0].runs[0]
                            font = first_run.font
                            # Clear and re-add with style
                            tf.clear()
                            new_p = tf.paragraphs[0]
                            new_run = new_p.add_run()
                            new_run.text = new_text
                            new_run.font.name = font.name
                            new_run.font.size = font.size
                            new_run.font.bold = font.bold
                            new_run.font.italic = font.italic
                            
                            # Safely handle color
                            if hasattr(font, 'color') and font.color is not None:
                                try:
                                    new_run.font.color.rgb = font.color.rgb
                                except (AttributeError, TypeError):
                                    # If color doesn't have RGB property, use a default color
                                    from pptx.dml.color import RGBColor
                                    new_run.font.color.rgb = RGBColor(0, 0, 0)  # Default to black
                        else:
                            shape.text = new_text  # fallback: plain replacement
            
            # Save the presentation
            prs.save(output_path)
            
            result = {"output_path": output_path, "status": "success"}
            
            # Store the interaction in memory
            self.memory.save_context(
                {"input": f"Create PowerPoint with template: {template_path}"},
                {"output": json.dumps(result)}
            )
            
            return result
            
        except Exception as e:
            return {"error": f"Failed to create PowerPoint: {str(e)}", "status": "error"}
    
    def _arun(self, template_path: str, output_path: str = "output/AI_Based_Pitch_Deck.pptx"):
        raise NotImplementedError("Async version not implemented")

class PitchDeckAgent:
    """Agent that orchestrates the pitch deck creation process"""
    
    def __init__(self):
        self.tools = [
            PowerPointStructureTool(),
            StructureAnalysisTool(),
            ContentGenerationTool(),
            PowerPointCreatorTool()
        ]
        
        self.prompt = PromptTemplate.from_template(
            """You are an expert pitch deck creator. Your task is to create a professional pitch deck using the available tools.
            
            Follow these steps:
            1. Extract the structure from the PowerPoint template using the extract_ppt_structure tool
            2. Analyze the structure using the analyze_structure tool to understand slide purposes
            3. Generate appropriate content using the generate_content tool based on the topic and analysis
            4. Create the final PowerPoint presentation using the create_powerpoint tool
            
            When using the tools:
            - For extract_ppt_structure: Pass the template path as a string
            - For analyze_structure: Pass the structure from the previous step
            - For generate_content: Pass the topic, structure, and analysis from previous steps
            - For create_powerpoint: Pass the template path, content, and output path
            
            If any step fails:
            - Check the error message in the response
            - Try to fix the issue based on the error
            - If the same error occurs three times, stop and report the issue
            
            You have access to these tools:
            {tools}
            
            Use the following format:
            Question: the input question you must answer
            Thought: you should always think about what to do
            Action: the action to take, should be one of [{tool_names}]
            Action Input: the input to the action
            Observation: the result of the action
            ... (this Thought/Action/Action Input/Observation can repeat N times)
            Thought: I now know the final answer
            Final Answer: the final answer to the original input question
            
            Question: {input}
            Thought: {agent_scratchpad}"""
        )
        
        self.agent = create_react_agent(
            llm=model,
            tools=self.tools,
            prompt=self.prompt
        )
        
        self.agent_executor = AgentExecutor(
            agent=self.agent,
            tools=self.tools,
            verbose=True,
            handle_parsing_errors=True,
        )
    
    def create_pitch_deck(self, template_path: str, topic: str, output_path: str = "output/AI_Based_Pitch_Deck.pptx") -> Dict:
        """
        Create a pitch deck using the agent system.
        
        Args:
            template_path (str): Path to the PowerPoint template
            topic (str): Detailed topic information
            output_path (str): Path to save the final PowerPoint
            
        Returns:
            Dict: Result of the pitch deck creation process
        """
        try:
            print("\n=== Starting Pitch Deck Creation Process ===")
            
            # Validate inputs
            if not os.path.exists(template_path):
                return {"status": "error", "error": f"Template file not found: {template_path}"}
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Get tools
            ppt_tool = self.tools[0]  # PowerPointStructureTool
            analysis_tool = self.tools[1]  # StructureAnalysisTool
            content_tool = self.tools[2]  # ContentGenerationTool
            powerpoint_tool = self.tools[3]  # PowerPointCreatorTool
            
            print("\n1. Extracting PowerPoint Structure...")
            # Extract structure
            structure_result = ppt_tool._run(template_path)
            if "error" in structure_result:
                print(f"Error extracting structure: {structure_result['error']}")
                return {"status": "error", "error": structure_result["error"]}
            structure = structure_result["structure"]
            print("Structure extracted successfully!")
            print("Structure:", structure)
            
            print("\n2. Analyzing Structure...")
            # Analyze structure
            analysis_result = analysis_tool._run(structure)
            if "error" in analysis_result:
                print(f"Error analyzing structure: {analysis_result['error']}")
                return {"status": "error", "error": analysis_result["error"]}
            analysis = analysis_result["analysis"]
            print("Structure analyzed successfully!")
            print("Analysis:", analysis)
            
            print("\n3. Generating Content...")
            # Generate content
            content_result = content_tool._run(topic, structure, analysis)
            if "error" in content_result:
                print(f"Error generating content: {content_result['error']}")
                return {"status": "error", "error": content_result["error"]}
            content = content_result["content"]
            print("Content generated successfully!")
            print("Content:", content)
            
            print("\n4. Creating PowerPoint...")
            # Set the content in the PowerPointCreatorTool
            powerpoint_tool.content_tool = content_tool
            
            # Create PowerPoint
            powerpoint_result = powerpoint_tool._run(template_path, output_path)
            if "error" in powerpoint_result:
                print(f"Error creating PowerPoint: {powerpoint_result['error']}")
                return {"status": "error", "error": powerpoint_result["error"]}
            print("PowerPoint created successfully!")
            
            print("\n=== Pitch Deck Creation Completed ===")
            return {"status": "success", "output_path": output_path}
            
        except Exception as e:
            print(f"\nError in pitch deck creation: {str(e)}")
            return {"status": "error", "error": str(e)}

def main():
    # Example usage
    template_path = "template/Black Elegant and Modern Startup Pitch Deck Presentation.pptx"
    output_path = "output/Black Elegant and Modern Startup Pitch Deck Presentation.pptx"
    
    # Detailed topic with company information
    topic = {
        "company": "TechFlow AI",
        "industry": "AI-Powered Business Process Automation",
        "founded": 2023,
        "location": "San Francisco, CA",
        "problem": [
            "Businesses waste 20+ hours per week on repetitive tasks",
            "Manual data entry leads to 15% error rate",
            "Companies lose $50B annually due to inefficient processes",
            "70% of employees report burnout from repetitive work"
        ],
        "solution": [
            "AI-powered workflow automation platform",
            "Reduces manual work by 80%",
            "99.9% accuracy in data processing",
            "Integrates with existing business tools",
            "Customizable for different industries"
        ],
        "target_market": [
            "Mid-size enterprises (100-1000 employees)",
            "Focus on finance, healthcare, and retail sectors",
            "$2B market size in target segments",
            "25% year-over-year market growth"
        ],
        "current_traction": [
            "50+ enterprise clients",
            "$2M ARR",
            "95% customer satisfaction",
            "40% month-over-month growth"
        ],
        "revenue_model": [
            "Subscription-based pricing",
            "Tiered plans: Basic ($499/mo), Pro ($999/mo), Enterprise (Custom)",
            "Additional revenue from customization services",
            "80% gross margin"
        ],
        "funding_need": [
            "Seeking $5M Series A",
            "18-month runway",
            "Product development: 40%",
            "Market expansion: 30%",
            "Team growth: 20%",
            "Operations: 10%"
        ]
    }
    
    # Create pitch deck using the agent system
    agent = PitchDeckAgent()
    result = agent.create_pitch_deck(template_path, json.dumps(topic), output_path)
    
    if result["status"] == "success":
        print(f"\nPitch deck created successfully at: {result['output_path']}")
    else:
        print("\nFailed to create pitch deck")
        print("Error:", result["error"])

if __name__ == "__main__":
    main() 