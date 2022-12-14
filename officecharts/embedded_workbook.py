import pandas
import io


def create(data: pandas.DataFrame) -> bytes:
    """Returns bytes array containing Excel workbook
    DataFrame is copied into Sheet1

    Args:
        data (pandas.DataFrame): DataFrame to embed

    Returns:
        bytes: Bytes array representing embedded Excel workbook
    """

    memory_stream = io.BytesIO()
    writer = pandas.ExcelWriter(memory_stream, engine='xlsxwriter', engine_kwargs={'options': {'in_memory': True}})
    data.to_excel(writer, 'Sheet1')
    writer.close()

    return memory_stream.getvalue()

