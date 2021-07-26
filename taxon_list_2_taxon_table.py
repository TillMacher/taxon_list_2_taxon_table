import pandas as pd
import string, os, subprocess, requests_html, json, re, sys
import os.path
from os import path
from pathlib import Path
import PySimpleGUI as sg

def create_reference_table(raw_taxa_list):

    # import morpho table
    df = pd.read_excel(raw_taxa_list)
    try:
        raw_taxa_list = df["Raw taxa"].values.tolist()
    except:
        sg.PopupError("Cannot find the raw taxa column!")
        raise RuntimeError

    ##################################################################################################################
    # download the taxonomy

    def get_gbif(query):

        gbif_taxonomy_list = []

        ## create an html session
        with requests_html.HTMLSession() as session:

            ## request that name
            r = session.get('https://api.gbif.org/v1/species?name=%s&limit=1' % query)

            ## parse json
            results = json.loads(r.text)

            taxonomic_ranks = ["phylum", "class", "order", "family", "genus"]

            for rank in taxonomic_ranks:
                try:
                    gbif_taxonomy_list.append(results['results'][0][rank])
                except:
                    gbif_taxonomy_list.append("NA")

        return gbif_taxonomy_list

    taxonomy_dict = {}

    for taxon in raw_taxa_list:
        # collect clear species names
        if len(taxon.split(" ")) == 2:
            taxonomy = get_gbif(taxon)
            taxonomy_dict[taxon] = taxonomy + [taxon]
        # collect higher taxa
        elif len(taxon.split(" ")) == 1:
            taxonomy = get_gbif(taxon)
            taxonomy_dict[taxon] = taxonomy + ["NA"]
        # collect ambigous names
        else:
            taxonomy = get_gbif(taxon.split(" ")[0])
            taxonomy_dict[taxon] = taxonomy + ["NA"]

    # convert to dataframe
    df_out = pd.DataFrame.from_dict(taxonomy_dict, orient='index')
    # rename column
    df_out = df_out.rename(columns={0: "Phylum", 1: "Class", 2: "Order", 3: "Family", 4: "Genus", 5: "Species"})
    # color all flags
    def color_negative_red(val):
        color = 'red' if val == "NA" else 'black'
        return 'color: %s' % color
    df_out = df_out.style.applymap(color_negative_red)

    with pd.ExcelWriter("reference_table.xlsx") as writer:
        df_out.to_excel(writer, sheet_name='reference_table')
        formatted_taxon_df.to_excel(writer, sheet_name='formatted_taxon_list', index=False)

    sg.Popup("A new reference table has been build. Please manually check for missing information.")

def update_reference_table(raw_taxa_list, reference_table):
    # load the reference raw_taxa_list as daraframe
    try:
        df = pd.read_excel(raw_taxa_list)
        raw_taxa_list_new = df["Raw taxa"].values.tolist()
    except:
        sg.PopupError("Cannot find the raw taxa column in the raw taxa list!")
        sys.exit()
    # load the reference table as daraframe
    try:
        reference_table_df = pd.read_excel(reference_table, sheet_name='reference_table')
    except:
        sg.PopupError("Cannot find reference table sheet")
        sys.exit()
    try:
        raw_taxa_list_old = reference_table_df["Raw taxa"].values.tolist()
    except:
        sg.PopupError("Cannot find the raw taxa column in the reference table!")
        sys.exit()

    taxa_to_update_list = list(set(raw_taxa_list_new) - set(raw_taxa_list_old))

    def get_gbif(query):

        gbif_taxonomy_list = []

        ## create an html session
        with requests_html.HTMLSession() as session:

            ## request that name
            r = session.get('https://api.gbif.org/v1/species?name=%s&limit=1' % query)

            ## parse json
            results = json.loads(r.text)

            taxonomic_ranks = ["phylum", "class", "order", "family", "genus"]

            for rank in taxonomic_ranks:
                try:
                    gbif_taxonomy_list.append(results['results'][0][rank])
                except:
                    gbif_taxonomy_list.append("NA")

        return gbif_taxonomy_list

    taxonomy_dict = {}

    for taxon in taxa_to_update_list:
        # collect clear species names
        if len(taxon.split(" ")) == 2:
            taxonomy = get_gbif(taxon)
            taxonomy_dict[taxon] = taxonomy + [taxon]
        # collect higher taxa
        elif len(taxon.split(" ")) == 1:
            taxonomy = get_gbif(taxon)
            taxonomy_dict[taxon] = taxonomy + ["NA"]
        # collect ambigous names
        else:
            taxonomy = get_gbif(taxon.split(" ")[0])
            taxonomy_dict[taxon] = taxonomy + ["NA"]

    # convert to dataframe
    df_out = pd.DataFrame.from_dict(taxonomy_dict, orient='index')
    # rename column
    df_out = df_out.rename(columns={0: "Phylum", 1: "Class", 2: "Order", 3: "Family", 4: "Genus", 5: "Species"})
    # rename index
    df_out.index.names = ['Raw taxa']
    # reset index
    df_out = df_out.reset_index()
    # merge with existing reference table
    df_out = reference_table_df.append(df_out)
    # reset the reset
    df_out = df_out.reset_index().drop(['index'], axis=1)
    # color all flags
    def color_negative_red(val):
        color = 'red' if val == "NA" else 'black'
        return 'color: %s' % color
    df_out = df_out.style.applymap(color_negative_red)

    with pd.ExcelWriter(reference_table) as writer:
        df_out.to_excel(writer, sheet_name='reference_table', index=False)

    sg.Popup("The reference table has been updated. Please manually check for missing information.")

def convert_table_format(conversion_table,reference_table, save_as):

    # load the morphology table as dataframe
    conversion_table_df = pd.read_excel(Path(conversion_table))
    # extract the sites ("Geäwsser") and the taxa ("Taxon") from the table
    try:
        conversion_table_taxa_per_site = conversion_table_df[["Site", "Taxa"]].values.tolist()
        conversion_table_sites = list(set(conversion_table_df["Site"].values.tolist()))
    except:
        sg.PopupError("Cannot find the columns 'Site' and 'Taxa'!")
        raise RuntimeError

    # load the reference table as daraframe
    try:
        reference_table_df = pd.read_excel(Path(reference_table), sheet_name='reference_table')
    except:
        sg.PopupError("Cannot find reference table sheet")
        raise RuntimeError
    # convert the reference table to a dictionary, where the key is the offical taxon and the values is the downloaded taxonomy
    reference_table_taxa_dict = {}
    for taxon in reference_table_df.values.tolist():
        reference_table_taxa_dict[taxon[0]] = taxon[1:]

    # create a dict with the present taxa per site
    presence_dict = {}
    for site in conversion_table_sites:
        for site_taxon in conversion_table_taxa_per_site:
            if site in site_taxon:
                if site not in presence_dict.keys():
                    presence_dict[site] = [site_taxon[1]]
                else:
                    presence_dict[site] = presence_dict[site] + [site_taxon[1]]

    # build the presence/absence table equivalently to the read table
    # store the table as a dict first
    morpho_table_list = []
    i = 1
    # iterate through all available taxa from the reference table
    for taxon, taxonomy in reference_table_taxa_dict.items():
        # store the taxonomy and the presence absence data per site in a list
        sites_presence_absence_list = []
        # check for each site if the taxon is present or not
        for site in conversion_table_sites:
            if taxon in presence_dict[site]:
                sites_presence_absence_list.append(1)
            else:
                sites_presence_absence_list.append(0)
        # save the taxonomy if the taxon was present in at least one site
        if max(sites_presence_absence_list) != 0:
            # add an OTU identifier
            OTU = "OTU_" + str(i)
            morpho_table_list.append([OTU] + taxonomy + ["100", "Morphology", "GCAT"] + sites_presence_absence_list)
            i += 1

    # create a dataframe
    df_out = pd.DataFrame(morpho_table_list)
    # remove spaces from site names!
    conversion_table_sites = [site.replace(" ", "_") for site in conversion_table_sites]
    # adjust the column names
    df_out.columns = ["IDs", "Phylum", "Class", "Order", "Family", "Genus", "Species", "Similarity", "Status", "seq"] + conversion_table_sites
    # export the data frame
    df_out.to_excel(save_as, index=False, sheet_name="TaXon table")
    # finish the script
    sg.Popup("Finished table conversion.")

def convert_matrix_format(conversion_table,reference_table, save_as):

    # load the morphology table as dataframe
    conversion_table_df = pd.read_excel(Path(conversion_table))
    conversion_table_df = conversion_table_df.fillna(0)

    if conversion_table_df.columns.tolist()[0] != "Taxa":
        sg.PopupError("Cannot find the column 'Taxa'!")
        raise RuntimeError

    # load the reference table as dataframe
    reference_table_df = pd.read_excel(Path(reference_table), sheet_name='reference_table')
    # convert the reference table to a dictionary, where the key is the offical taxon and the values is the downloaded taxonomy
    reference_table_taxa_dict = {}
    for taxon in reference_table_df.values.tolist():
        reference_table_taxa_dict[taxon[0]] = taxon[1:]

    # build the presence/absence table equivalently to the read table
    # store the table as a dict first
    morpho_table_list = []
    for i, entry in enumerate(conversion_table_df.values.tolist()):
        taxon = entry[0]
        abundances = entry[1:]
        OTU = ["OTU_" + str(i+1)]
        taxonomy = reference_table_taxa_dict[taxon]
        morpho_table_list.append(OTU + taxonomy+ ["100", "Morphology", "GCAT"] + abundances)

    # create a dataframe
    df_out = pd.DataFrame(morpho_table_list)
    # remove spaces from site names!
    conversion_table_sites = [site.replace(" ", "_") for site in conversion_table_df.columns[1:].tolist()]
    # adjust the column names
    df_out.columns = ["IDs", "Phylum", "Class", "Order", "Family", "Genus", "Species", "Similarity", "Status", "seq"] + conversion_table_sites
    # export the data frame
    df_out.to_excel(save_as, index=False, sheet_name="TaXon table")
    # finish the script
    sg.Popup("Finished table conversion.")

##################################################################################################################
##################################################################################################################
##################################################################################################################
##################################################################################################################
##################################################################################################################
# create a window

sg.theme('Reddit')

layout = [[sg.Text('Table converter', font=('Arial', 12, "bold"))],
          [sg.Text("")],
          [sg.Text("Raw taxa list:", size=(15,1)), sg.Input(), sg.FileBrowse(key='raw_taxa_list')],
          [sg.Text("Rerefence table:", size=(15,1)), sg.Input(), sg.FileBrowse(key='reference_table')],
          [sg.Text("", size=(15,1)), sg.Button("Create new", key="run_create_reference_table"), sg.Button("Update existing", key="run_update_reference_table")],
          [sg.Text("",size=(15,2))],
          [sg.Text("Input table:", size=(15,1)), sg.Input(), sg.FileBrowse(key='conversion_table')],
          [sg.Text("Table format:", size=(15,1)), sg.Radio("Table", "format", key="table_format", default=True), sg.Radio("Matrix", "format", key="matrix_format")],
          [sg.Text("Save as:", size=(15,1)), sg.Input(), sg.SaveAs(key="save_as")],
          [sg.Text("Conversion", size=(15,1)), sg.Button('Run', key="convert_morpho_table")],
          [sg.Text("")],
          [sg.Exit()]]

window = sg.Window('Morpho table converter', layout)

while True:                             # The Event Loop
    try:

        event, values = window.read()
        raw_taxa_list = values["raw_taxa_list"]
        reference_table = values["reference_table"]
        save_as = values["save_as"]
        conversion_table = values["conversion_table"]

        if event in (None, 'Exit'):
            break

        if event == "run_create_reference_table":
            if raw_taxa_list == "":
                sg.PopupError("Please provide a raw taxa list")
            else:
                warning = sg.PopupOKCancel("Warning: This will overwrite any existing reference table!")
                if warning == "OK":
                    create_reference_table(raw_taxa_list)

        if event == "run_update_reference_table":
            if raw_taxa_list == "":
                sg.PopupError("Please provide a raw taxa list")
            elif reference_table == "":
                sg.PopupError("Please provide a reference table")
            else:
                update_reference_table(raw_taxa_list, reference_table)

        if event == "convert_morpho_table":
            if conversion_table == "":
                sg.PopupError("Please provide an input table")
            elif reference_table == "":
                sg.PopupError("Please provide a reference table")
            elif save_as == "":
                sg.PopupError("Please provide an output file")
            else:
                if values["table_format"] == True:
                    convert_table_format(conversion_table, reference_table, save_as)
                elif values["matrix_format"] == True:
                    convert_matrix_format(conversion_table, reference_table, save_as)

    except RuntimeError:
        print("")

window.close()
