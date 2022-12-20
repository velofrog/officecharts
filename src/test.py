import officecharts

officecharts.create_linechart(officecharts.data, title="My Chart", axis_y_title="Cumulative")

officecharts.create_linechart(officecharts.data, title="My Chart", axis_y_title="Cumulative",
                              label_endpoints=True,
                              theme=officecharts.Theme(font=officecharts.Font(family="Times New Roman"),
                                                       plot_area=officecharts.Style(fill_colour="#cacaca20",
                                                                                    colour="#cacaca"))
                              )
