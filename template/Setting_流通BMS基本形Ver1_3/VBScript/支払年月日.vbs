function EmptyOrDateTime(in1, out1)
	on error resume next

    if len(in1) = 10 then
	out1 = mid(in1, 3, 2) + mid(in1, 6, 2) + mid(in1, 9, 2)
    end if

	EmptyOrDateTime = 0
end function
