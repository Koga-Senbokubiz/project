function EmptyOrDateTime(in1, out1)
	on error resume next

    if len(in1) = 6 then
        out1 = "20" + mid(in1, 1, 2) + "-" + mid(in1, 3, 2) + "-" + mid(in1, 5)
    end if

	EmptyOrDateTime = 0
end function
