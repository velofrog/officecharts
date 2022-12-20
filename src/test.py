import officecharts

officecharts.create_linechart(officecharts.data, title="My Chart", axis_y_title="Cumulative")

officecharts.create_linechart(officecharts.data, title="My Chart", axis_y_title="Cumulative",
                              label_endpoints=True,
                              theme=officecharts.Theme(font=officecharts.Font(family="Trebuchet MS"),
                                                       legend_position=None,
                                                       axis_y_format="0.00%",
                                                       grid_major_x=officecharts.Style(width=0.75),
                                                       plot_area=officecharts.Style(fill_colour="#cacaca20",
                                                                                    colour="#cacaca"))
                              )
