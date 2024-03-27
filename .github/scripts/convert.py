import pandas as pd
import json
import os
from tabulate import tabulate
import plotly.graph_objects as go

def ConvertExcelToMD(PathToExcel):
    print("Converting Excel to MD...")
    onto_list = pd.ExcelFile(PathToExcel).sheet_names
    # Remove Sheet names with non-ontology names:
    ignore_sheets = ['Template mit Beispiel', 'List_zu_betrachtende_Ontologien']
    onto_list = [sheet for sheet in onto_list if sheet not in ignore_sheets]
    
    for onto_name in onto_list:
        print(f"Processing ontology: {onto_name}")
        table = pd.read_excel(PathToExcel, sheet_name=onto_name)
        
        with open("./json/GeneralStructure.json") as f: 
            ontodata_dict = json.load(f)
        
        data_table = ["Ontology"]
        data_table.extend(list(table.Ontology.dropna(how='all')))
        
        for superkey in ontodata_dict:
            if superkey == "Comments":
                list_index = data_table.index(superkey)
                ontodata_dict[superkey] = data_table[list_index+1:]
            else:
                for key in ontodata_dict[superkey]:        
                    list_index = data_table.index(key)
                    ontodata_dict[superkey][key] = data_table[list_index+1]
        
        with open(f"./json/{onto_name}.json", "w") as f:
            json.dump(ontodata_dict, f)
            
        with open("./json/md-translator.json") as f:     
            translator_dict = json.load(f)
            
        outstring = "## " + onto_name + " - " + ontodata_dict["Ontology"]["Ontology Name"] + "\n\n"   
        # DomainRadarPlotter(onto_name)  # Uncomment this if you want to generate radar plots
        
        outstring += f"\n ## Radarplot \n\n [HTML-Version](../radarplots/Radarplot_{onto_name}.html) ![Radarplot for Domains of ontology {onto_name}](../radarplots/Radarplot_{onto_name}.svg) \n"
        
        table_string = "|Aspect |Description| \n |:---|:---|\n"
        
        for key in translator_dict:
            if key == "Comments":
                outstring += "## Comments\n\n"
                for i in ontodata_dict[key]:
                    outstring += str(i) + "\n"
            else:
                outstring += f"## {key}\n\n"
                outstring += table_string
                for dict_list in translator_dict[key]:
                    outstring += f"| {list(dict_list.values())[0]} | {ontodata_dict[key][list(dict_list.keys())[0]]} |\n"
                outstring += "\n"
                
        with open(f'./ontology_metadata/{onto_name}.md', 'w') as f:
            f.write(outstring)
    print("Conversion completed.")

def load_ontologies_metadata():
    print("Loading ontologies metadata...")
    json_list = [f for f in os.listdir('./json/') if (f.endswith('.json') and f != "GeneralStructure.json" and f!= "md-translator.json" and f!= "ontology_domains.json")]
    metadata_dict = {}
    for json_name in json_list:
        with open(f'./json/{json_name}') as file:
            onto_metadata = json.load(file)
            metadata_dict[onto_metadata["Ontology"]["Ontology Acronym"]] = onto_metadata
    print("Metadata loaded.")
    return metadata_dict        

def UpdateMainReadme(): 
    print("Updating Main Readme...")
    path = './ontology_metadata/'
    markdown_list = [s for s in os.listdir(path) if s.endswith('.md')]
    print_list = "| Link to Markdown | Ontology Name |\n |:---:|:---|\n"
    
    for i in markdown_list:
        ontology_name = i.replace('.md','')
        with open(f'./json/{ontology_name}.json') as dictFile:
            onto_dict = json.load(dictFile)
        
        print_list += f'| [{ontology_name}] | {onto_dict["Ontology"]["Ontology Name"]} |\n'
    
    print_list += '\n'
    for i in markdown_list:
        print_list += f'[{i.replace(".md","")}]: ./ontology_metadata/{i}\n' 
    
    with open('./Main_Readme_Update.txt', 'w') as f:
        f.write(print_list)
   
    md_dict = load_ontologies_metadata()
    key_dom_interest = "Domain of Interest Represented (contained, related: broader/narrower, missing)"
    domains_of_interest = list(md_dict[list(md_dict.keys())[0]][key_dom_interest].keys())
    domain_dict = {}

    for domain in domains_of_interest:
        onto_list = []
        for onto_abbrev in md_dict:
            dict_entry = md_dict[onto_abbrev][key_dom_interest][domain]
            if ("contained" in dict_entry) or ("related:narrower" in dict_entry.replace(" ","")):
                onto_list.append(onto_abbrev) 
        domain_dict[domain] = onto_list
    
    DomainRadarPlotter_all_ontologies()

    with open('./Main_Readme_Update.txt', 'a') as f:
        f.write("\n## Map of Ontologies for Catalysis Research Domains\n\n")
        f.write("\n The ontologies are classified with regards to their research domain [here](./Radarplots.md).\n")
        f.write("\n [Here](./Radarplot.html) you can find the Radar plot as interactive plot.\n")
        f.write("\n ![Map of Ontologies for Catalysis Research Domains](./Fig2-OntoMap.svg)\n")
    print("Main Readme updated.")

def Heatmap_to_Markdown():
    print("Converting Heatmap to Markdown...")
    df = pd.read_excel('MappingHeatmap.xlsx')
    md_dict = load_ontologies_metadata()
    df = df.fillna('')
    for row in df.index:
        for col in df.columns:
            cell_value = df.at[row, col]
            if type(cell_value) != str and cell_value != df.at[row,'Unnamed: 0']:
                new_value = f'[' + str(int(cell_value)) +'](/mapping/'+col+'_'+df.at[row,'Unnamed: 0'] +'.md)'
                df.at[row, col] = new_value
            
            if df.at[row,'Unnamed: 0'] == col:
                df.at[row, col] = md_dict[col]["Ontology Characteristics"]["Class count"]
                
    for col in df.columns:
        df.rename(columns={col: '['+col+']'}, inplace=True)
    
    for row in df.index:
        repl_str = '['+df.at[row,'[Unnamed: 0]']+']'
        df.at[row,'[Unnamed: 0]'] = repl_str
        
    df.rename(columns={'[Unnamed: 0]': ''}, inplace=True)
    
    df.to_markdown('Heatmap_Classes.md', index=False)
    print("Heatmap conversion to Markdown completed.")

def Mappings_to_Markdown():
    print("Converting Mappings to Markdown...")
    file_list = [f for f in os.listdir('./mapping/') if f.endswith('.xlsx')]
    
    for file in file_list:
        df = pd.read_excel('./mapping/'+file)
        del df['Unnamed: 0']
        df.to_markdown('mapping/' + file.replace('.xlsx', '.md'), index=False)
    print("Mappings conversion to Markdown completed.")

def DomainRadarPlotter_all_ontologies():
    print("Generating Radar plots for all ontologies...")
    md_dict = load_ontologies_metadata()
    key_dom_interest = "Domain of Interest Represented (contained, related: broader/narrower, missing)"
    domains_of_interest = list(md_dict[list(md_dict.keys())[0]][key_dom_interest].keys())
    
    domain_dict_c = {}
    domain_dict_c_n = {}
    domain_dict_c_n_b = {}

    for domain in domains_of_interest:
        onto_list_c = []
        onto_list_c_n = []
        onto_list_c_n_b = []
        
        for onto_abbrev in md_dict:
            dict_entry = md_dict[onto_abbrev][key_dom_interest][domain]
            
            if "contained" in dict_entry:
                onto_list_c.append(onto_abbrev) 
                
            if "contained" in dict_entry or "related:narrower" in dict_entry.replace(" ",""):
                onto_list_c_n.append(onto_abbrev) 
                
            if "contained" in dict_entry or "related:narrower" in dict_entry.replace(" ","") or "related:broader" in dict_entry.replace(" ",""):
                onto_list_c_n_b.append(onto_abbrev) 
        
        domain_dict_c[domain] = onto_list_c
        domain_dict_c_n[domain] = onto_list_c_n
        domain_dict_c_n_b[domain] = onto_list_c_n_b
    
    plotlist_c = [len(domain_dict_c[i]) for i in domains_of_interest]
    plotlist_c_n = [len(domain_dict_c_n[i]) for i in domains_of_interest]
    plotlist_c_n_b = [len(domain_dict_c_n_b[i]) for i in domains_of_interest]
    
    plotlist_c.extend([plotlist_c[0]])
    plotlist_c_n.extend([plotlist_c_n[0]])
    plotlist_c_n_b.extend([plotlist_c_n_b[0]])
    domains_of_interest.extend([domains_of_interest[0]])
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatterpolar(
          r= plotlist_c_n_b,
          theta=domains_of_interest,
          fill='toself',
          marker=dict(color='lightcoral'),
          name='related: broader'
    ))
      
    fig.add_trace(go.Scatterpolar(
          r=plotlist_c_n,
          theta=domains_of_interest,
          fill='toself',
          marker=dict(color='gold'),
          name='related: narrower'
    ))

    fig.add_trace(go.Scatterpolar(
          r=plotlist_c,
          theta=domains_of_interest,
          fill='toself',
          marker=dict(color='seagreen'),
          name='contained'
    ))
      
    fig.update_layout(
       autosize=False,
       width=1000,
       height=800,
       polar=dict(
        radialaxis=dict(
          visible=True,
          range=[0, max(plotlist_c_n_b)]
        )),
       legend=dict(
           yanchor="bottom",
           y=-0.2,
           xanchor="center",
           x=0.5
           ),
      showlegend=True
    )
    
    fig.write_html("Radarplot.html")
    fig.write_image("Radarplot.svg")
    print("Radar plots generation completed.")

def DomainRadarPlotter(ontology_name):
    print(f"Generating Radar plot for {ontology_name}...")
    md_dict = load_ontologies_metadata()
    key_dom_interest = "Domain of Interest Represented (contained, related: broader/narrower, missing)"
    domains_of_interest = list(md_dict[ontology_name][key_dom_interest].keys())
    
    domain_dict_c = {}
    domain_dict_c_n = {}
    domain_dict_c_n_b = {}

    for domain in domains_of_interest:
        onto_list_c = []
        onto_list_c_n = []
        onto_list_c_n_b = []
        
        dict_entry = md_dict[ontology_name][key_dom_interest][domain]
        
        if "contained" in dict_entry:
            onto_list_c = 1
        else:
            onto_list_c = 0
            
        if "contained" in dict_entry or "related:narrower" in dict_entry.replace(" ",""):
            onto_list_c_n = 2
        else:
            onto_list_c_n = 0
            
        if "contained" in dict_entry or "related:narrower" in dict_entry.replace(" ","") or "related:broader" in dict_entry.replace(" ",""):
            onto_list_c_n_b = 3
        else:
            onto_list_c_n_b = 0
    
        domain_dict_c[domain] = onto_list_c
        domain_dict_c_n[domain] = onto_list_c_n
        domain_dict_c_n_b[domain] = onto_list_c_n_b
    
    plotlist_c = [domain_dict_c[i] for i in domains_of_interest]
    plotlist_c_n = [domain_dict_c_n[i] for i in domains_of_interest]
    plotlist_c_n_b = [domain_dict_c_n_b[i] for i in domains_of_interest]
    
    plotlist_c.extend([plotlist_c[0]])
    plotlist_c_n.extend([plotlist_c_n[0]])
    plotlist_c_n_b.extend([plotlist_c_n_b[0]])
    domains_of_interest.extend([domains_of_interest[0]])
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatterpolar(
          r=plotlist_c_n_b,
          theta=domains_of_interest,
          fill='toself',
          marker=dict(color='lightcoral'),
          name='3 = related: broader'
    ))
       
    fig.add_trace(go.Scatterpolar(
          r=plotlist_c_n_b,
        theta=domains_of_interest,
          fill='toself',
          marker=dict(color='gold'),
          name='2 = related: narrower'
    ))

    fig.add_trace(go.Scatterpolar(
          r= plotlist_c,
          theta=domains_of_interest,
          fill='toself',
          marker=dict(color='seagreen'),
          name='1 = contained'
    ))
    
    fig.update_layout(
      autosize=False,
      width=1000,
      height=800,
      polar=dict(
        radialaxis=dict(
          visible=True,
          range=[0, 3]
        )),
       legend=dict(
           yanchor="bottom",
           y=-0.2,
           xanchor="center",
           x=0.5
           ),
      showlegend=True
    )
    
    fig.write_html(f"./radarplots/Radarplot_{ontology_name}.html")
    fig.write_image(f"./radarplots/Radarplot_{ontology_name}.svg")
    print(f"Radar plot for {ontology_name} generated.")

def run():    
    print("Starting the run function")
    Master_Table = './master_table/MT_OntoWorldMap_2023-10-11.xlsx'
    try:
        ConvertExcelToMD(Master_Table)
        print("ConvertExcelToMD function completed")
    except Exception as e:
        print("Error occurred in ConvertExcelToMD function:", e)
    
    try:
        UpdateMainReadme()
        print("UpdateMainReadme function completed")
    except Exception as e:
        print("Error occurred in UpdateMainReadme function:", e)

if __name__ == "__main__":
    run()
