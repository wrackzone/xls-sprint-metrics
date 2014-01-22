require "date"

class Metric
    
    def initialize (name, description, field, aggregation, day)  
      @name        = name  
      @description = description
      @field       = field
      @aggregation = aggregation
      @day         = day
    end 
       
    def name= name
        @name = name
    end
    
    def description= description
        @description = description
    end
    
    def field= field
        @field = field
    end
    
    def aggregation= aggregation   # :sum, :count, :max, :min, :mean etc.
        @aggregation = aggregation
    end
       
    def day= day # :first, :last etc.
        @day = day
    end
    
    def day_in_snapshot snapshot,day

		if snapshot["_ValidTo"][0..3] == "9999"
			if day >= Date.parse(snapshot["_ValidFrom"])
				return true
			end
		end
		return day < Date.parse(snapshot["_ValidTo"]) && day >= Date.parse(snapshot["_ValidFrom"])

	end

    
    def calculate( iteration, snapshots )
        d = @day == :first ? Date.parse(iteration.StartDate) + 1 : Date.parse(iteration.EndDate) + 1
        sfd = (snapshots.collect { |snapshot| snapshot if day_in_snapshot(snapshot,d)}).compact!
        
        # group by object ID
        groups = sfd.group_by { |sn| sn["ObjectID"] }
        
        if groups.keys.size() == 0
            return 0
        end
        
        if @aggregation == :count
            return groups.size()
        end
        if @aggregation == :sum
            value = 
            (groups.keys.map { |group| groups[group].last[@field] ? groups[group].last[@field] : 0 }).reduce(0,:+)
            return value
        end
        
    end
    
end

class Day1TaskEstimate < Metric
    def initialize
        super("Task Estimate First Day (Hours)", "Total estimate of tasks for first day of iteration","TaskEstimateTotal",:sum,:first)
    end
end

class Day1TaskCount < Metric
    def initialize
        super("Tasks on First Day (Count)", "","TaskEstimateTotal",:count,:first)
    end
end

class Day2TaskCount < Metric
    def initialize
        super("Tasks on Last Day (Count)", "Total estimate of tasks for first day of iteration","TaskEstimateTotal",:count,:last)
    end
end

