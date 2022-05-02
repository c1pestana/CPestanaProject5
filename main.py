##CAITLIN PESTANA
##NOT DONE!!!! show_percentage_change_map() does not run


import openpyxl
import plotly.graph_objects
import openpyxl.utils
from state_abbrev import us_state_to_abbrev
import numbers

def main():
    pop_worksheet = open_worksheet()
    show_pop_change = should_display_pop_change()
    if show_pop_change is True:
        return show_pop_change_map(pop_worksheet)
    else:
        return show_percentage_change_map(pop_worksheet)

def open_worksheet():
    income_excel = openpyxl.load_workbook("countyPopChange2020-2021.xlsx")
    data_sheet = income_excel.active
    return data_sheet

def should_display_pop_change():
    response = input("Should we display a map of total population changes?")
    if response.lower() == "yes":
        return True
    else:
        return False

def show_pop_change_map(pop_worksheet):

    list_of_state_abbrev = []
    population_changes = []
    pop_worksheet.delete_rows(0)
    for row in pop_worksheet.rows:
        if row[4].value == 0 :
            state_cell = row[5]
            first_cell = row[11]
            population_changes.append(int(first_cell.value))

            state_name = state_cell.value
            state_abbrev = us_state_to_abbrev[state_name]
            list_of_state_abbrev.append(state_abbrev)

    map_to_show = plotly.graph_objects.Figure(
        data=plotly.graph_objects.Choropleth(
            locations= list_of_state_abbrev,
            z=population_changes,
            locationmode="USA-states",
            colorscale='Picnic',
            colorbar_title="population change 2021"))

    map_to_show.update_layout(
        title_text="Population Change 2021",
        geo_scope="usa")

    map_to_show.show()

def show_percentage_change_map(pop_worksheet):
    list_of_state_abbrev = []
    list_of_pop_change_percent = []
    for row in pop_worksheet.rows:
        first_cell = row[5]
        population_change_cell = row[11]
        pop_estimate_cell = row[9]
        population_change_value = population_change_cell.value
        pop_estimate_value = pop_estimate_cell.value
        state_name = first_cell.value
        if not isinstance(population_change_value, numbers.Number):
            continue
        if not isinstance(pop_estimate_value, numbers.Number):
            continue
        if state_name not in us_state_to_abbrev:
            continue
        state_abbrev = us_state_to_abbrev[state_name]
        list_of_state_abbrev.append(state_abbrev)
        npopchg2021_cell_number = openpyxl.utils.cell.column_index_from_string('n') - 1
        popestimate2021_cell_number = openpyxl.utils.cell.column_index_from_string('b') - 1
        npopchg2021_cell = row[npopchg2021_cell_number]
        popestimate2021_cell = row[popestimate2021_cell_number]
        npopchg2021 = npopchg2021_cell.value
        popestimate2021 = popestimate2021_cell.value
        pop_change_percent = npopchg2021 / popestimate2021
        pop_change_percent = pop_change_percent *100
        list_of_pop_change_percent.append(pop_change_percent)


    map_to_show = plotly.graph_objects.Figure(
        data=plotly.graph_objects.Choropleth(
            locations=list_of_state_abbrev,
            z= list_of_pop_change_percent,
            locationmode="USA-states",
            colorscale='Picnic',
            colorbar_title='percentage of population change'
        )
    )
    map_to_show.update_layout(
        title_text="Percentage change of population 2020-2021",
        geo_scope="usa"
    )
    map_to_show.show()


main()
