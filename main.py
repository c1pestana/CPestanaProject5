import openpyxl
import plotly
from state_abbrev import us_state_to_abbrev
import numbers

def open_worksheet():
    income_excel = openpyxl.load_workbook("countyPopChange2020-2021.xlsx")
    pop_worksheet = income_excel.active
    return pop_worksheet

def main():
    pop_worksheet = open_worksheet()
    show_pop_change = should_display_pop_change()
    if show_pop_change is True:
        return show_pop_change_map(pop_worksheet)
    else:
        return show_percentage_change_map(pop_worksheet)

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
            locations= list_of_state_abbrev[:],
            z=population_changes,
            locationmode="USA-states",
            colorscale='Picnic',
            colorbar_title="population change 2021"))

        map_to_show.update_layout(
            title_text="Population Change 2021",
            geo_scope="usa")

        map_to_show.show()

def show_percentage_change_map(pop_worksheet):
    # answer = int(input("What is the cutoff for income growth"))
    # print(f"States Whose Median Income Grew more than {answer} in 2015-2020")
    list_of_state_abbrev = []
    list_of_income_changes = []
    for row in pop_worksheet.rows:
        first_cell = row[0]
        income_cell = row[1]
        income_value = income_cell.value
        if not isinstance(income_value, numbers.Number):
            continue
        if first_cell.value not in us_state_to_abbrev:
            continue
        state_name = first_cell.value
        state_abrev = us_state_to_abbrev[state_name]
        list_of_state_abbrev.append(state_abrev)
        income2015_cell_number = openpyxl.utils.cell.column_index_from_string('n') - 1
        income2015_cell = row[income2015_cell_number]
        income2015 = income2015_cell.value
        income_change = income_value - income2015
        list_of_income_changes.append(income_change)
        # income_chance_percent = income_change/income_value
        # income_chance_percent = income_chance_percent *100

    map_to_show = plotly.graph_objects.Figure(
        data=plotly.graph_objects.Choropleth(
            locations=us_state_to_abbrev,
            z=list_of_income_changes,
            locationmode="USA-states",
            colorscale='Picnic',
            colorbar_title="amount of income"
        )
    )
    map_to_show.update_layout(
        title_text="Income Change 2015 to 2020",
        geo_scope="usa"
    )
    map_to_show.show()


main()

