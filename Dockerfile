FROM mono

ADD . /src

WORKDIR /src

RUN ls -lrt /src

RUN ls -lrt /src/Data

RUN mono --help

#RUN mono nuget.exe restore

RUN msbuild SiqiSecondPptMerge.sln /p:Configuration=Release

#RUN mono SiqiSecondPptMerge/bin/Debug/SiqiSecondPptMerge.exe

CMD [ "sh", "-c", "mono SiqiSecondPptMerge/bin/Release/SiqiSecondPptMerge.exe && ls -lrt /src/Data" ]
