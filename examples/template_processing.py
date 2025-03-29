#!/usr/bin/env python3
"""
Template processing example for llamadocx.

This example demonstrates advanced template processing capabilities,
including nested fields, repeating sections, and conditional content.
"""

import os
import json
from pathlib import Path
from llamadocx import Document, Template
from llamadocx.metadata import set_author, set_title


def create_report_template():
    """Create a complex report template with various field types."""
    print("Creating a report template...")
    
    # Create a new document
    doc = Document()
    
    # Add company logo placeholder
    doc.add_paragraph("[COMPANY LOGO]")
    
    # Add title page
    doc.add_heading("{{report.title}}", level=1)
    doc.add_paragraph("Prepared for: {{client.name}}")
    doc.add_paragraph("Prepared by: {{author.name}}")
    doc.add_paragraph("Date: {{report.date}}")
    
    # Add table of contents header
    doc.add_heading("Table of Contents", level=1)
    
    # Add repeating section for table of contents
    doc.add_paragraph("{{#sections}}")
    doc.add_paragraph("{{number}}. {{title}} ................... {{page}}")
    doc.add_paragraph("{{/sections}}")
    
    # Add executive summary
    doc.add_heading("Executive Summary", level=1)
    doc.add_paragraph("{{report.summary}}")
    
    # Add introduction
    doc.add_heading("1. Introduction", level=1)
    doc.add_paragraph("{{report.introduction}}")
    
    # Add company overview section
    doc.add_heading("2. Company Overview", level=1)
    doc.add_paragraph("{{client.description}}")
    
    # Add company details in a table
    table = doc.add_table(rows=4, cols=2)
    table.rows[0].cells[0].text = "Company Name"
    table.rows[0].cells[1].text = "{{client.name}}"
    table.rows[1].cells[0].text = "Industry"
    table.rows[1].cells[1].text = "{{client.industry}}"
    table.rows[2].cells[0].text = "Founded"
    table.rows[2].cells[1].text = "{{client.founded}}"
    table.rows[3].cells[0].text = "Location"
    table.rows[3].cells[1].text = "{{client.location}}"
    
    # Add findings section
    doc.add_heading("3. Key Findings", level=1)
    doc.add_paragraph("{{report.findings_intro}}")
    
    # Add repeating section for findings
    doc.add_paragraph("{{#findings}}")
    doc.add_heading("3.{{index}}. {{title}}", level=2)
    doc.add_paragraph("{{description}}")
    
    # Add a table for metrics if available
    doc.add_paragraph("{{#has_metrics}}")
    doc.add_paragraph("Metrics:")
    metrics_table = doc.add_table(rows=1, cols=3)
    metrics_table.rows[0].cells[0].text = "Metric"
    metrics_table.rows[0].cells[1].text = "Value"
    metrics_table.rows[0].cells[2].text = "Benchmark"
    doc.add_paragraph("{{#metrics}}")
    doc.add_paragraph("{{name}}: {{value}} (Benchmark: {{benchmark}})")
    doc.add_paragraph("{{/metrics}}")
    doc.add_paragraph("{{/has_metrics}}")
    
    doc.add_paragraph("{{/findings}}")
    
    # Add recommendations section
    doc.add_heading("4. Recommendations", level=1)
    doc.add_paragraph("Based on our findings, we recommend the following actions:")
    
    # Add repeating section for recommendations
    doc.add_paragraph("{{#recommendations}}")
    doc.add_paragraph("{{index}}. {{text}}")
    doc.add_paragraph("Priority: {{priority}} | Timeline: {{timeline}}")
    doc.add_paragraph("{{/recommendations}}")
    
    # Add conclusion
    doc.add_heading("5. Conclusion", level=1)
    doc.add_paragraph("{{report.conclusion}}")
    
    # Add contact information
    doc.add_heading("Contact Information", level=1)
    doc.add_paragraph("For questions regarding this report, please contact:")
    doc.add_paragraph("Name: {{author.name}}")
    doc.add_paragraph("Email: {{author.email}}")
    doc.add_paragraph("Phone: {{author.phone}}")
    
    # Add metadata
    set_title(doc, "Report Template")
    set_author(doc, "LlamaSearch.AI")
    
    # Save the template
    template_path = Path("report_template.docx")
    doc.save(template_path)
    print(f"Template saved to: {template_path.absolute()}")
    
    return template_path


def create_sample_data():
    """Create sample data for the report template."""
    data = {
        "report": {
            "title": "2024 Business Analysis Report",
            "date": "March 31, 2024",
            "summary": "This executive summary provides an overview of our analysis of TechCorp's current business operations, challenges, and opportunities. Our assessment identified several key areas for improvement and growth opportunities.",
            "introduction": "This report presents a comprehensive analysis of TechCorp's business operations, market position, and strategic opportunities. The analysis was conducted over a six-week period and included interviews with key stakeholders, market research, and competitive analysis.",
            "findings_intro": "Our analysis identified the following key findings:",
            "conclusion": "TechCorp is well-positioned for growth but faces significant challenges in a rapidly evolving market. By implementing the recommendations outlined in this report, TechCorp can strengthen its market position, improve operational efficiency, and drive sustainable growth."
        },
        "client": {
            "name": "TechCorp, Inc.",
            "description": "TechCorp is a mid-sized technology company specializing in enterprise software solutions for the healthcare and finance industries. With over 500 employees across three offices, TechCorp has established a strong presence in the North American market.",
            "industry": "Enterprise Software",
            "founded": "2005",
            "location": "San Francisco, CA"
        },
        "author": {
            "name": "Jane Smith",
            "title": "Senior Business Analyst",
            "email": "jane.smith@example.com",
            "phone": "(555) 123-4567"
        },
        "sections": [
            {"number": "1", "title": "Introduction", "page": "3"},
            {"number": "2", "title": "Company Overview", "page": "5"},
            {"number": "3", "title": "Key Findings", "page": "8"},
            {"number": "4", "title": "Recommendations", "page": "15"},
            {"number": "5", "title": "Conclusion", "page": "20"}
        ],
        "findings": [
            {
                "index": "1",
                "title": "Market Position",
                "description": "TechCorp currently holds approximately 12% market share in the healthcare software segment, positioning it as the fourth-largest provider in this space. However, the company's market share has declined by 2% over the past two years due to increased competition from both established players and new entrants.",
                "has_metrics": True,
                "metrics": [
                    {"name": "Market Share", "value": "12%", "benchmark": "15%"},
                    {"name": "Year-over-Year Growth", "value": "-0.5%", "benchmark": "+3%"},
                    {"name": "Customer Satisfaction", "value": "78%", "benchmark": "82%"}
                ]
            },
            {
                "index": "2",
                "title": "Product Portfolio",
                "description": "TechCorp's product portfolio includes six major software solutions, with 70% of revenue coming from just two products. This concentration presents a significant risk, especially as these products are in mature market segments with limited growth potential.",
                "has_metrics": True,
                "metrics": [
                    {"name": "Product Concentration", "value": "70%", "benchmark": "<50%"},
                    {"name": "R&D Investment", "value": "8% of revenue", "benchmark": "12% of revenue"},
                    {"name": "New Product Revenue", "value": "5%", "benchmark": "15%"}
                ]
            },
            {
                "index": "3",
                "title": "Operational Efficiency",
                "description": "Our analysis of TechCorp's operations identified several inefficiencies in the development process, customer onboarding, and support services. These inefficiencies result in longer time-to-market for new features and higher-than-industry-average customer churn.",
                "has_metrics": False
            }
        ],
        "recommendations": [
            {
                "index": "1",
                "text": "Diversify product portfolio by investing in new solution development for emerging market segments",
                "priority": "High",
                "timeline": "12-18 months"
            },
            {
                "index": "2",
                "text": "Implement agile development methodologies across all product teams to improve time-to-market",
                "priority": "Medium",
                "timeline": "6-9 months"
            },
            {
                "index": "3",
                "text": "Develop a customer success program to improve retention and satisfaction metrics",
                "priority": "High",
                "timeline": "3-6 months"
            },
            {
                "index": "4",
                "text": "Explore strategic partnerships or acquisitions to rapidly expand into adjacent markets",
                "priority": "Medium",
                "timeline": "12-24 months"
            }
        ]
    }
    
    # Save data to JSON file for reference
    data_path = Path("report_data.json")
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    
    print(f"Sample data saved to: {data_path.absolute()}")
    
    return data


def process_template(template_path, data):
    """Process the template with the provided data."""
    print(f"\nProcessing template: {template_path}")
    
    # Create Template object
    template = Template(template_path)
    
    # Get template fields
    fields = template.get_fields()
    print(f"Template contains {len(fields)} fields")
    
    # Merge data into template
    template.merge_fields(data)
    
    # Save the processed document
    output_path = Path("generated_report.docx")
    template.save(output_path)
    
    print(f"Generated report saved to: {output_path.absolute()}")
    return output_path


if __name__ == "__main__":
    # Create a template
    template_path = create_report_template()
    
    # Create sample data
    data = create_sample_data()
    
    # Process the template
    process_template(template_path, data)
    
    print("\nTemplate processing example completed. Check the current directory for output files.") 