class resultObject:
    def __init__(self, surveys, locations, shots, observations, samples, creator, ID_Indices):
        self.surveys = surveys
        self.locations = locations
        self.shots = shots
        self.observations = observations
        self.samples = samples

        #Saves GlobalID values
        self.survey_GlobalID = self.surveys[ID_Indices[0]]
        self.site_GlobalID = self.locations[ID_Indices[1]]
        self.shot_GlobalID = self.shots[ID_Indices[2]]
        self.obs_GlobalID = self.observations[ID_Indices[3]]
        self.sample_GlobalID = self.samples[ID_Indices[4]]

        #Save creator from survey list
        self.creator = creator

        #Prepares final data list
        self.collation = []

    def collate(self, globalHeader):
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
        result = [None] * len(list)
        header_result = [None] * len(list)
        index = 0

        if len(template) < len(list):
            template += ([-1] * (len(list) - len(template)))
        

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