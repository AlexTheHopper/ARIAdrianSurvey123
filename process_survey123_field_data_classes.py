

class SiteSections:
    def __init__(self, so_site_id, so_section_number):
        self.site_id = so_site_id
        self.section_number = so_section_number
        self.index = dict(
            {0: self.site_id,
             1: self.section_number}
        )

    def __getitem__(self, key):
        return self.index[key]

    def __setitem__(self, key, newvalue):
        self.index[key] = newvalue

class SiteObs:
    def __init__(self, so_site_id, so_section_number, so_species, so_collected, so_observed, so_collected2, so_shot_id, so_obs_id):
        self.site_id = so_site_id
        self.section_number = so_section_number
        self.species = so_species
        self.collected = so_collected
        self.observed = so_observed
        self.collected_left = so_collected2
        self.shot_id = so_shot_id
        self.obs_id = so_obs_id
        self.index = dict(
            {0: self.site_id,
             1: self.section_number,
             2: self.species,
             3: self.collected,
             4: self.observed,
             5: self.collected_left,
             6: self.shot_id,
             7: self.obs_id}
        )

    def __getitem__(self, key):
        return self.index[key]

    def __setitem__(self, key, newvalue):
        self.index[key] = newvalue
# #
# ##    def updateCollected(self, key, new_collected_total):
# #        self.__setitem__(key, new_collected_total)


class SiteSurvey:
    def __init__(self, ss_site_id, ss_site_code, ss_survey_date, ss_gear_type, ss_personnel1, ss_personnel2, ss_depth_secchi, ss_depth_max, ss_depth_avg, ss_section_condition, ss_time_start, ss_time_end, ss_project_name, ss_survey_notes, ss_water_qual_depth, ss_ec_25c, ss_water_temp, ss_do_mgl, ss_do_perc, ss_ph, ss_turbidity_ntu, ss_chlorophyll, ss_creationdate, ss_creator, ss_editdate, ss_editor, ss_data_x, ss_data_y, ss_x_start, ss_y_start, ss_x_finish, ss_y_finish, ss_shot_id, ss_section_number, ss_electro_seconds, ss_soak_minutes_per_unit, ss_section_time_start, ss_section_time_end, ss_volts, ss_amps, ss_pulses_per_second, ss_percent_duty_cycle):
        self.site_id = ss_site_id
        self.site_code = ss_site_code
        self.survey_date = ss_survey_date
        self.gear_type = ss_gear_type
        self.personnel1 = ss_personnel1
        self.personnel2 = ss_personnel2
        self.depth_secchi = ss_depth_secchi
        self.depth_max = ss_depth_max
        self.depth_avg = ss_depth_avg
        self.section_condition = ss_section_condition
        self.time_start = ss_time_start
        self.time_end = ss_time_end
        self.project_name = ss_project_name
        self.survey_notes = ss_survey_notes
        self.water_qual_depth = ss_water_qual_depth
        self.ec_25c = ss_ec_25c
        self.water_temp = ss_water_temp
        self.do_mgl = ss_do_mgl
        self.do_perc = ss_do_perc
        self.ph = ss_ph
        self.turbidity_ntu = ss_turbidity_ntu
        self.chlorophyll = ss_chlorophyll
        self.creationdate = ss_creationdate
        self.creator = ss_creator
        self.editdate = ss_editdate
        self.editor = ss_editor
        self.data_x = ss_data_x
        self.data_y = ss_data_y
        self.x_start = ss_x_start
        self.y_start = ss_y_start
        self.x_finish = ss_x_finish
        self.y_finish = ss_y_finish
        self.shot_id = ss_shot_id
        self.section_number = ss_section_number
        self.electro_seconds = ss_electro_seconds
        self.soak_minutes_per_unit = ss_soak_minutes_per_unit
        self.section_time_start = ss_section_time_start
        self.section_time_end = ss_section_time_end
        self.volts = ss_volts
        self.amps = ss_amps
        self.pulses_per_second = ss_pulses_per_second
        self.percent_duty_cycle = ss_percent_duty_cycle
        self.index = dict(
               {'k_site_id': self.site_id,
                'k_site_code': self.site_code,
                'k_survey_date': self.survey_date,
                'k_gear_type': self.gear_type,
                'k_personnel1': self.personnel1,
                'k_personnel2': self.personnel2,
                'k_depth_secchi': self.depth_secchi,
                'k_depth_max': self.depth_max,
                'k_depth_avg': self.depth_avg,
                'k_section_condition': self.section_condition,
                'k_time_start': self.time_start,
                'k_time_end': self.time_end,
                'k_project_name': self.project_name,
                'k_survey_notes': self.survey_notes,
                'k_water_qual_depth': self.water_qual_depth,
                'k_ec_25c': self.ec_25c,
                'k_water_temp': self.water_temp,
                'k_do_mgl': self.do_mgl,
                'k_do_perc': self.do_perc,
                'k_ph': self.ph,
                'k_turbidity_ntu': self.turbidity_ntu,
                'k_chlorophyll': self.chlorophyll,
                'k_creationdate': self.creationdate,
                'k_creator': self.creator,
                'k_editdate': self.editdate,
                'k_editor': self.editor,
                'k_data_x': self.data_x,
                'k_data_y': self.data_y,
                'k_x_start': self.x_start,
                'k_y_start': self.y_start,
                'k_x_finish': self.x_finish,
                'k_y_finish': self.y_finish,
                'k_shot_id': self.shot_id,
                'k_section_number': self.section_number,
                'k_electro_seconds': self.electro_seconds,
                'k_soak_minutes_per_unit': self.soak_minutes_per_unit,
                'k_section_time_start': self.section_time_start,
                'k_section_time_end': self.section_time_end,
                'k_volts': self.volts,
                'k_amps': self.amps,
                'k_pulses_per_second': self.pulses_per_second,
                'k_percent_duty_cycle': self.percent_duty_cycle}
        )

    def __getitem__(self, key):
        return self.index[key]

    def __setitem__(self, key, newvalue):
        self.index[key] = newvalue
