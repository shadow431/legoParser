#FROM python:3-slim
FROM public.ecr.aws/lambda/python:3.8

COPY ./src/legoParser.py ${LAMBDA_TASK_ROOT}/app.py
COPY ./src/smartsheet.py ${LAMBDA_TASK_ROOT}

COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt --target "${LAMBDA_TASK_ROOT}"


CMD [ "app.handler" ]
