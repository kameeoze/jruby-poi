module POI
  class Area
    def initialize reference, version
      @ref = reference
      @version = version
    end

    def in workbook
      if single_cell_reference?
        ref = area.all_referenced_cells.first
        return [workbook.single_cell(ref)]
      end

      begin
        by_column = {}
        # refs = area.all_referenced_cells
        # slices = refs.enum_slice(refs.length/15)
        # slices.collect do |cell_refs|
        #   Thread.start do
        #     cell_refs.each do |cell_ref|
        #       first = workbook.worksheets[cell_ref.sheet_name].first_row
        #       last  = workbook.worksheets[cell_ref.sheet_name].last_row
        #       next unless cell_ref.row >= first && cell_ref.row <= last
        #
        #       # ref = POI::CELL_REF.new(c.format_as_string)
        #       cell = workbook.single_cell cell_ref
        #       (by_column[cell_ref.cell_ref_parts.collect.last] ||= []) << cell
        #     end
        #   end
        # end.each {|t| t.join}

        area.all_referenced_cells.each do |cell_ref|
          first = workbook.worksheets[cell_ref.sheet_name].first_row
          last  = workbook.worksheets[cell_ref.sheet_name].last_row
          next unless cell_ref.row >= first && cell_ref.row <= last

          # ref = POI::CELL_REF.new(c.format_as_string)
          cell = workbook.single_cell cell_ref
          (by_column[cell_ref.cell_ref_parts.to_a.last] ||= []) << cell
        end

        by_column.each do |key, cells|
          by_column[key] = cells.compact
        end

        if by_column.length == 1
          by_column.values.flatten
        else
          by_column
        end
      rescue
        []
      end
    end

    def single_cell_reference?
      area.single_cell? #@ref == getFirstCell.formatAsString rescue false
    end

    private
      def area
        @area ||= new_area_reference
      end

      def new_area_reference
        begin
          return column_reference
        rescue => x
          return POI::AREA_REF.new(@ref, @version)
        rescue
          raise "could not determine area reference for #{@ref}: #{$!.message}"
        end
      end

      def column_reference
        sheet_parts = @ref.split('!')
        area_parts  = sheet_parts.last.split(':')
        area_start  = "#{sheet_parts.first}!#{area_parts.first}"
        area_end    = area_parts.last

        POI::AREA_REF.getWholeColumn(@version, area_start, area_end)
      end
  end
end
