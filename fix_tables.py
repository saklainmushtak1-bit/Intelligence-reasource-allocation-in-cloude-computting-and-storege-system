import re

with open('generate_docx_report.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if 'format_table(t_sdg,' in line and 'sdg_table_data' not in line:
        new_lines.append('    format_table(t_sdg, [1.8, 2.3, 1.6], ["Goal", "UN Target Definition", "NextDrive Concrete Contribution"], sdg_table_data, header_bg="1B365D")\n')
        skip = True
        continue
    if 'format_table(t_lit,' in line and 'lit_data' not in line:
        new_lines.append('    format_table(t_lit, [1.3, 1.6, 1.3, 1.5], ["Author & Year", "Method / Focus", "Scaling Strategy", "Identified Gaps vs. NextDrive"], lit_data, header_bg="2C3E50")\n')
        skip = True
        continue
    if 'format_table(t_tools,' in line and 'tools_data' not in line:
        new_lines.append('    format_table(t_tools, [1.2, 0.9, 1.4, 2.2], ["Tool / Technology", "Version", "Role in NextDrive", "Engineering Justification"], tools_data, header_bg="1B365D")\n')
        skip = True
        continue
    if 'format_table(t_files,' in line and 'files_data' not in line:
        new_lines.append('    format_table(t_files, [1.4, 1.8, 2.5], ["Filename", "Location in Workspace", "Functional Role & Implementation Purpose"], files_data, header_bg="2C3E50")\n')
        skip = True
        continue
    if 'format_table(t_test,' in line and 'test_cases_data' not in line:
        new_lines.append('    format_table(t_test, [0.7, 1.4, 1.6, 1.4, 0.6], ["Test ID", "Test Scenario", "Input / Action", "Expected & Observed Outcome", "Result"], test_cases_data, header_bg="1B365D")\n')
        skip = True
        continue
    if 'format_table(t_bench,' in line and 'bench_data' not in line:
        new_lines.append('    format_table(t_bench, [1.0, 0.8, 0.7, 0.7, 0.9, 0.7, 0.9], ["Architecture", "Compute Allocation", "Power Draw", "Avg Util.", "Spike Latency", "SLA Breaches", "Efficiency Assessment"], bench_data, header_bg="2C3E50")\n')
        skip = True
        continue
    if 'format_table(t_team,' in line and 'team_data' not in line:
        new_lines.append('    format_table(t_team, [1.2, 1.1, 1.5, 1.9], ["Team Member", "Roll Number", "Primary Role & Focus", "Detailed Contributions & Technical Ownership"], team_data, header_bg="1B365D")\n')
        skip = True
        continue

    if skip:
        if line.strip().startswith('add_p(') or line.strip().startswith('t_') or line.strip().startswith('add_h'):
            skip = False
            new_lines.append(line)
        else:
            continue
    else:
        new_lines.append(line)

with open('generate_docx_report.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("generate_docx_report.py cleaned up successfully.")
