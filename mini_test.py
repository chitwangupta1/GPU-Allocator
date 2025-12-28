"""
Simple test to verify CrewAI + Gemini setup works
"""

import os
from dotenv import load_dotenv
from crewai import Agent, Task, Crew, LLM

print("="*60)
print("Testing CrewAI + Gemini Integration")
print("="*60)

# Load environment
load_dotenv()
api_key = os.getenv('GOOGLE_API_KEY')

if not api_key:
    print("❌ ERROR: GOOGLE_API_KEY not found")
    print("Please add it to your .env file")
    exit(1)

print(f"\n1. API Key loaded: {api_key[:6]}...{api_key[-4:]} ({len(api_key)} chars)")

# Initialize LLM
print("\n2. Initializing Gemini LLM...")
try:
    llm = LLM(
        model="gemini/gemini-2.0-flash-exp",
        api_key=api_key,
        temperature=0.7
    )
    print("   ✓ LLM initialized")
except Exception as e:
    print(f"   ❌ LLM initialization failed: {e}")
    exit(1)

# Create a simple agent
print("\n3. Creating test agent...")
try:
    test_agent = Agent(
        role='Test Agent',
        goal='Test if the setup works',
        backstory='A simple test agent to verify CrewAI works with Gemini',
        verbose=True,
        llm=llm
    )
    print("   ✓ Agent created")
except Exception as e:
    print(f"   ❌ Agent creation failed: {e}")
    exit(1)

# Create a simple task
print("\n4. Creating test task...")
try:
    test_task = Task(
        description='Say "Hello, GPU Marketplace!" and confirm you are working.',
        agent=test_agent,
        expected_output='A greeting message confirming the system works'
    )
    print("   ✓ Task created")
except Exception as e:
    print(f"   ❌ Task creation failed: {e}")
    exit(1)

# Create crew and run
print("\n5. Running test crew...")
try:
    crew = Crew(
        agents=[test_agent],
        tasks=[test_task],
        verbose=True
    )
    print("   ✓ Crew created")
    
    print("\n" + "="*60)
    print("EXECUTING TEST...")
    print("="*60)
    
    result = crew.kickoff()
    
    print("\n" + "="*60)
    print("TEST RESULT:")
    print("="*60)
    print(result)
    print("\n" + "="*60)
    print("✓✓✓ ALL TESTS PASSED! ✓✓✓")
    print("Your CrewAI + Gemini setup is working correctly!")
    print("="*60)
    
except Exception as e:
    print(f"\n❌ Crew execution failed: {e}")
    import traceback
    traceback.print_exc()
    exit(1)