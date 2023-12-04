#Object for raw data rows:
class resultObject:
    def __init__(self, surveys, locations, shots, observations, samples, creator):
        self.surveys = surveys
        self.locations = locations
        self.shots = shots
        self.observations = observations
        self.samples = samples

        #Saves GlobalID values
        self.survey_GlobalID = surveys[0]
        self.site_GlobalID = locations[0]
        self.shot_GlobalID = shots[0]
        self.obs_GlobalID = observations[0]
        self.sample_GlobalID = samples[0]

        #Save creator from survey list
        self.creator = creator

        #Prepares final data list
        self.collation = []

    def collate(self, globalHeader):
        #Collates all data into one single list for output.
        self.collation += self.surveys
        self.collation += self.locations
        self.collation += self.shots
        self.collation += self.observations
        self.collation += self.samples

        self.collation += [self.survey_GlobalID]
        self.collation += [self.site_GlobalID]
        self.collation += [self.shot_GlobalID]
        self.collation += [self.obs_GlobalID]
        self.collation += [self.sample_GlobalID]
        self.collation += [self.creator]

        #Combine species columns:
        species = self.collation[globalHeader.index('species_obs')] if self.collation[globalHeader.index('species_samp')] is None else self.collation[globalHeader.index('species_samp')]
        self.collation[globalHeader.index('species_obs')] = species
        self.collation.pop(globalHeader.index('species_samp'))

    def order(self, list, template, header, allocation):
        #Orders data according to templates at top of main.
        result = [None] * len(list)
        header_result = [None] * len(list)
        index = 0

        #Ensure template and data have same length
        if len(template) < len(list):
            print('template,list',len(template),len(list))
            diff = (len(list) - len(template))
            template += ([-1] * diff)
            print('template,list',len(template),len(list))
        elif len(template) > len(list):
            print('template,list',len(template),len(list))
            diff = (len(template) - len(list))
            del template[-diff:]
            print('template,list',len(template),len(list))
            
        

        while index < (len(list)):
            extra = 0
            
            #Joins two cells (eg personnel 1&2)
            if template[index] == 'j':
                try:
                    instruction = str(list[index]) + ', ' + str(list[index + 1])
                    result[template[index+1]] = instruction

                    header_name = str(header[index]) + ', ' + str(header[index+1])
                    header_result[template[index+1]] = header_name
                    header_result.pop(-1)
                    result.pop(-1)
                    extra = 1
                except:
                    print("ERROR: JOIN ATTEMPT AT LAST INDEX")

            #Ignore cells
            elif template[index] < 0:
               
                header_result.pop(-1)
                result.pop(-1)
               

            #Move cells
            elif int(template[index]) >= 0:
                
                instruction = round(template[index])

                result[instruction] = list[index]
                header_result[instruction] = header[index]

            index += 1
            index += extra

        #Send reordered correct list
        if allocation == 'survey':
            self.surveys = result
        elif allocation == 'location':
            self.locations = result
        elif allocation == 'shot':
            self.shots = result
        elif allocation == 'obs':
            self.observations = result
        elif allocation == 'sample':
            self.samples = result
        return header_result