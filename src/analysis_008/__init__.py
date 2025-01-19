from src.utils import flatten, queryGQL
import pandas as pd

#####################################################################################
#
# https://github.com/Hujnis/OUISproject
#
#####################################################################################
query = """query ($where: EventInputFilter){
  result: eventPage(where:$where, limit: 1000){
    id
    name
    place
    placeId
    startdate
    enddate
    eventType{
      id
      name
    }
    groups{
      id
      name
    }
  }
}"""

async def resolve_json(variables, cookies):
    assert "where" in variables, f"missing where in parameters"
    jsonresponse = await queryGQL(
        query=query,
        variables=variables,
        cookies=cookies
        )
    
    data = jsonresponse.get("data", {"result": None})
    result = data.get("result", None)
    assert result is not None, f"got {jsonresponse}"

    return result

async def resolve_flat_json(variables, cookies):
    jsonData = await resolve_json(variables=variables, cookies=cookies)
    mapper = {
        "eventID": "id",
        "eventName": "name",
        "placeID": "placeId",
        "place" :"place",
        "startDate": "startdate",
        "endDate": "enddate",
        "eventTypeID": "eventType.id",
        "eventType": "eventType.name",
        "groupID": "groups.id",
        "groupName": "groups.name",
    }
    # print(jsonData, flush=True)
    pivotdata = list(flatten(jsonData, {}, mapper))
    return pivotdata

async def resolve_df_pivot(variables, cookies):
    pivotdata = await resolve_flat_json(variables=variables, cookies=cookies)

    # print(pivotdata)
    df = pd.DataFrame(pivotdata)

    pdf = pd.pivot_table(df, values="user_fullname", index="classification_sem", columns=["classification_level"], aggfunc="count")

    return pdf


#####################################################################################
#
# 
#
#####################################################################################
import string
import openpyxl
from tempfile import NamedTemporaryFile
from fastapi import APIRouter, Request, Query, Response
from ..utils import process_df_as_html_page
import json
import re
import io
import datetime

def createRouter(prefix):
    mainpath = "/facilities/events"
    tags = ["Hodiny na učebnách"]

    router = APIRouter(prefix=prefix)
    WhereDescription = "filtr omezující vybrané skupiny"
    @router.get(f"{mainpath}/table", tags=tags, summary="HTML tabulka s daty pro výpočet kontingenční tabulky")
    async def user_classification_html(
        request: Request,
        where: str = Query(description=WhereDescription)
    ):
        "HTML tabulka s daty pro výpočet kontingenční tabulky"
        wherevalue = None if where is None else re.sub(r'{([^:"]*):', r'{"\1":', where) 
        wherejson = json.loads(wherevalue)
        pd = await resolve_flat_json(
            variables={
                "where": wherejson
            },
            cookies=request.cookies
        )
        df = pd.DataFrame(pd)
        return await process_df_as_html_page(df)
    
    @router.get(f"{mainpath}/flatjson", tags=tags, summary="Data ve formátu JSON transformována do podoby vstupu pro kontingenční tabulku")
    async def user_classification_flat_json(
        request: Request, 
        where: str = Query(description=WhereDescription), 
    ):
        "Data ve formátu JSON transformována do podoby vstupu pro kontingenční tabulku"
        wherevalue = None if where is None else re.sub(r'{([^:"]*):', r'{"\1":', where) 
        wherejson = json.loads(wherevalue)
        pd = await resolve_flat_json(
            variables={
                "where": wherejson
            },
            cookies=request.cookies
        )
        return pd

    @router.get(f"{mainpath}/json", tags=tags, summary="Data ve formátu JSON (stromová struktura) nevhodná pro kontingenční tabulku")
    async def user_classification_json(
        request: Request, 
        where: str = Query(description=WhereDescription), 
        startdate: datetime.datetime = Query(description=""),
        enddate: datetime.datetime = Query(description="")
    ):
        "Data ve formátu JSON (stromová struktura) nevhodná pro kontingenční tabulku"
        wherevalue = None if where is None else re.sub(r'{([^:"]*):', r'{"\1":', where) 
        wherejson = json.loads(wherevalue)
        pd = await resolve_json(
            variables={
                "where": wherejson,
                "startdate": f"(startdate)",
            },
            cookies=request.cookies
        )
        return pd

    @router.get(f"{mainpath}/xlsx", tags=tags, summary="Xlsx soubor doplněný o data v záložce 'data' (podle xlsx vzoru)")
    async def user_classification_xlsx(
        request: Request, 
        where: str = Query(description=WhereDescription), 
    ):
        "Xlsx soubor doplněný o data v záložce 'data' (podle xlsx vzoru)"
        wherevalue = None if where is None else re.sub(r'{([^:"]*):', r'{"\1":', where) 
        wherejson = json.loads(wherevalue)
        flat_json = await resolve_flat_json(
            variables={
                "where": wherejson
            },
            cookies=request.cookies
        )

        with open('./src/xlsx/vzor2.xlsx', 'rb') as f:
            content = f.read()
        
        memory = io.BytesIO(content)
        resultFile = openpyxl.load_workbook(filename=memory)
        
        resultFileData = resultFile['data']
        
        for (rid, item) in enumerate(flat_json):
            for col, value in zip(string.ascii_uppercase, item.values()):
                cellname = f"{col}{rid+2}"
                resultFileData[cellname] = value

        with NamedTemporaryFile() as tmp:
            # resultFile.save(tmp.name)
            resultFile.save(tmp)
            tmp.seek(0)
            stream = tmp.read()
            headers = {
                'Content-Disposition': 'attachment; filename="Analyza.xlsx"'
            }
            return Response(stream, media_type='application/vnd.ms-excel', headers=headers)
        
    return router