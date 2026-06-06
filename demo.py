#!/usr/bin/env python3
"""
ClimateLenz Align & Paste Tool - Demo Script
Shows how to use the tool programmatically.
"""

from app import MaterialProcessor, TemplateAligner

# Sample CSRD report material
sample_material = """1. Executive Summary

ClimateLenz has prepared this sustainability report in accordance with the Corporate Sustainability Reporting Directive (CSRD) and European Sustainability Reporting Standards (ESRS).

The report covers the fiscal year 2025 and includes disclosures on environmental, social, and governance matters.

2. Scope and Methodology

Our reporting approach follows the double materiality principle as mandated by ESRS 1 General Requirements.

Key standards applied:
- ESRS E1 Climate Change
- ESRS E2 Pollution
- ESRS S1 Own Workforce
- ESRS G1 Business Conduct

Data collection methods:
- Supplier surveys and questionnaires
- Energy consumption records
- Employee demographic data
- Governance policy reviews

3. Environmental Disclosures

3.1 Climate Change (ESRS E1)

The company has identified climate-related risks across its value chain. Transition risks include regulatory changes and market shifts. Physical risks include supply chain disruptions from extreme weather events.

Key metrics:
- Scope 1 emissions: 1,250 tCO2e
- Scope 2 emissions: 890 tCO2e
- Scope 3 emissions: 12,400 tCO2e

3.2 Pollution (ESRS E2)

Air emissions are monitored at all manufacturing facilities. Water discharge quality meets EU standards. Waste generation has decreased 15% year-over-year.

4. Social Disclosures

4.1 Own Workforce (ESRS S1)

Total employees: 1,240
Gender distribution: 52% male, 47% female, 1% non-binary
Training hours per employee: 24 hours annually
Employee satisfaction score: 7.8/10

4.2 Workers in the Value Chain

Supplier audits conducted: 45
Critical findings: 3
Remediation completed: 100%

5. Governance Disclosures

5.1 Business Conduct (ESRS G1)

Anti-corruption policies are in place across all operations. Whistleblower mechanism active since 2023. Zero confirmed incidents of corruption in reporting period.

6. Conclusion and Next Steps

The company is committed to continuous improvement in sustainability reporting. Next reporting period will expand coverage to include biodiversity (ESRS E4) and circular economy (ESRS E5) disclosures.
"""

# Process the material
processor = MaterialProcessor(sample_material)
aligner = TemplateAligner(processor)

# Show detected sections
print("=" * 60)
print("CLIMATELENZ ALIGN & PASTE TOOL - DEMO")
print("=" * 60)
print(f"\nDetected {len(processor.sections)} sections:")
for i, section in enumerate(processor.sections, 1):
    print(f"  {i}. {section['title']} ({len(section['content'])} lines)")

# Generate Word document
print("\n" + "=" * 60)
print("Generating Word document...")
doc = aligner.align_to_word()
if doc:
    doc.save("demo_output.docx")
    print("✓ Saved: demo_output.docx")

# Generate Excel workbook
print("\nGenerating Excel workbook...")
wb = aligner.align_to_excel()
if wb:
    wb.save("demo_output.xlsx")
    print("✓ Saved: demo_output.xlsx")

# Generate PowerPoint
print("\nGenerating PowerPoint presentation...")
prs = aligner.align_to_pptx()
if prs:
    prs.save("demo_output.pptx")
    print("✓ Saved: demo_output.pptx")

print("\n" + "=" * 60)
print("Demo complete! Check the generated files.")
print("=" * 60)