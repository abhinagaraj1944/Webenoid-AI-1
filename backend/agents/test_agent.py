from dotenv import load_dotenv
load_dotenv()

from agents.excel_agent import ExcelAgent

def main():
    print("Starting ExcelAgent standalone test...\n")
    agent = ExcelAgent()

    test_inputs = [
        "sum",
        "average",
        "max",
        "hello",
        "calculate percentage",
        "how many values are in column A?",
        "foobar"
    ]

    for prompt in test_inputs:
        result = agent.analyze(prompt)
        print(f"Input: {prompt}")
        print("Agent response:", result)
        print("-" * 40)

if __name__ == "__main__":
    main()
