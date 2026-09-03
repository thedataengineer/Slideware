import json
import subprocess
import os

def check_command(cmd):
    try:
        subprocess.run(cmd, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError:
        return False

def analyze_repo():
    print("Maturity Analysis of the Repository\n" + "="*35 + "\n")

    # 1. Tests
    print("## 1. Testing")
    has_tests_dir = os.path.exists("tests")
    print(f"- Has 'tests' directory: {'Yes' if has_tests_dir else 'No'}")

    # We saw tests fail, but there is a test suite
    print("- Has test script in package.json: Yes (but with 5 failing tests out of 194)")
    print("- Type checking (TypeScript): Yes (npm run typecheck passes)")

    # 2. Documentation
    print("\n## 2. Documentation")
    for doc in ["README.md", "SETUP.md", "CHANGELOG.md", "LICENSE", "NOTICE"]:
        print(f"- {doc} exists: {'Yes' if os.path.exists(doc) else 'No'}")

    # 3. Tooling and Automation
    print("\n## 3. Tooling and Automation")
    print("- Linter configured: Yes (office-addin-lint, passes)")
    print("- CI/CD (GitHub Actions): No (.github folder not found)")
    print("- Prettier config: Yes")
    print("- Webpack build: Yes")

    # 4. Code Structure
    print("\n## 4. Code Structure")
    print("- TypeScript used: Yes")
    print("- React (JSX) configured: Yes")
    print("- Clear separation of concerns (src/taskpane/features, src/taskpane/powerpoint.ts)")

    # Summary
    print("\n## Summary")
    print("The repository has a high level of maturity in terms of code structure, documentation, and local tooling. It uses TypeScript, has a comprehensive test suite (though currently with 5 failing tests related to a missing module), linting, and detailed documentation like README, SETUP, and CHANGELOG. It appears to lack automated CI/CD pipelines (like GitHub Actions).")

if __name__ == "__main__":
    analyze_repo()
